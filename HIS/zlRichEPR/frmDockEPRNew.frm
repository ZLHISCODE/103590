VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDockEPRNew 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   0  'None
   Caption         =   "��������"
   ClientHeight    =   6135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3555
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   3555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin XtremeReportControl.ReportControl rptList 
      Height          =   1320
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   3330
      _Version        =   589884
      _ExtentX        =   5874
      _ExtentY        =   2328
      _StockProps     =   0
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      ShowHeader      =   0   'False
   End
   Begin VB.Frame frmBaby 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   3255
      Begin VB.ComboBox cboӤ�� 
         Height          =   300
         ItemData        =   "frmDockEPRNew.frx":0000
         Left            =   600
         List            =   "frmDockEPRNew.frx":0007
         Style           =   2  'Dropdown List
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   100
         Width           =   2475
      End
      Begin VB.Label lblӤ�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ĸӤ"
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   200
         Width           =   360
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   3255
      Begin VB.TextBox txtSearchKey 
         Height          =   300
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "�����������ƶ�λ"
         Top             =   105
         Width           =   3030
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgThis 
      Height          =   840
      Left            =   135
      TabIndex        =   2
      Top             =   5235
      Width           =   2775
      _cx             =   4895
      _cy             =   1482
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
      Cols            =   3
      FixedRows       =   1
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
      AutoResize      =   0   'False
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   555
      Top             =   105
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockEPRNew.frx":0015
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockEPRNew.frx":05AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockEPRNew.frx":0B49
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockEPRNew.frx":10E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockEPRNew.frx":167D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockEPRNew.frx":1C17
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockEPRNew.frx":21B1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   1470
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockEPRNew.frx":26C3
            Key             =   "ǩ��"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockEPRNew.frx":2A15
            Key             =   "���δ�ӡ"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockEPRNew.frx":2FAF
            Key             =   ""
            Object.Tag             =   "3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockEPRNew.frx":3549
            Key             =   ""
            Object.Tag             =   "99"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockEPRNew.frx":3AE3
            Key             =   ""
            Object.Tag             =   "90001"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockEPRNew.frx":407D
            Key             =   ""
            Object.Tag             =   "90002"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockEPRNew.frx":4617
            Key             =   ""
            Object.Tag             =   "90003"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockEPRNew.frx":49B1
            Key             =   ""
            Object.Tag             =   "90004"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockEPRNew.frx":4D4B
            Key             =   "ˢ��"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockEPRNew.frx":B5AD
            Key             =   "Ԥˢ��"
         EndProperty
      EndProperty
   End
   Begin VB.Line LineBottom 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   3225
      Y1              =   885
      Y2              =   885
   End
   Begin VB.Label lblNeaten 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����ʱ������"
      Height          =   180
      Left            =   0
      TabIndex        =   3
      Top             =   930
      Width           =   1350
      WordWrap        =   -1  'True
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   45
      Top             =   105
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin VB.Line LinTop 
      BorderColor     =   &H00808080&
      X1              =   150
      X2              =   3375
      Y1              =   585
      Y2              =   585
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ݲ��˵�������������������µĲ�����"
      Height          =   180
      Left            =   150
      TabIndex        =   0
      Top             =   105
      Width           =   3555
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmDockEPRNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const conColumn_ͼ�� = 0
Const conColumn_ID = 1
Const conColumn_�¼� = 2
Const conColumn_���� = 3
Private Enum mc
    ID = 1
    �¼� = 2
    ���� = 3
    ���� = 4
    ˵�� = 5
End Enum
'-----------------------------------------------------
'�����¼�
'-----------------------------------------------------
Public Event NewClick(ByVal FileId As Long, ByVal babyNum As Long) '��������²�����ť

'-----------------------------------------------------
'�������
Private mbytScene As Byte       'ʹ�ó���

Private mlngDeptId As Long
Private mlngPatiId As Long
Private mlngVisit As Long
Private mblnShowAll As Boolean
Private mstrPrivs As String
Private mlngAdviceID As Long    'ҽ��ID
Private mOldSearchKey As String  '�����ؼ���
Private mHeight As Long 'Ӥ��ѡ�������ʱ�����߶�
Private mlngSRow As Long

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
Dim intNum As Integer
    On Error Resume Next
    Select Case Control.ID
    Case conMenu_Edit_NewItem
        If rptList.FocusedRow Is Nothing Then Exit Sub
        If frmBaby.Visible Then
            intNum = Me.cboӤ��.ItemData(cboӤ��.ListIndex)
        End If
        RaiseEvent NewClick(rptList.FocusedRow.Record(conColumn_ID).Value, intNum)
    Case conMenu_View_Show
        rptList.PreviewMode = Not rptList.PreviewMode
        Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "�ļ�˵��", IIf(rptList.PreviewMode, 1, 0))
    Case conMenu_View_ShowAll
        mblnShowAll = Not mblnShowAll
        Call zlRefList(mbytScene, mlngPatiId, mlngVisit, mlngDeptId, mstrPrivs, mlngAdviceID)
        Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��ʾ����", IIf(mblnShowAll, 1, 0))
    End Select
End Sub
Private Sub cbsThis_Resize()
    Dim lngScaleLeft As Long, lngScaleTop As Long, lngScaleRight As Long, lngScaleBottom As Long
    Call Me.cbsThis.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
    
    Err = 0: On Error Resume Next
    If vfgThis.TextMatrix(1, 0) = "" Then
        vfgThis.Height = 0: vfgThis.Visible = False
        lblNeaten.Height = 0: lblNeaten.Visible = False
    Else
        vfgThis.Height = lngScaleBottom / 3: vfgThis.Visible = True
        lblNeaten.Height = 250: lblNeaten.Visible = True
    End If
    With Me.Frame
        .Left = 0: .Width = Me.ScaleWidth: .Top = 350
    End With
    With Me.lblNote
        .Left = 150: .Width = Me.ScaleWidth - .Left * 2
        .Top = 500 + 500
    End With
    With Me.frmBaby
        .Left = 0: .Width = Me.ScaleWidth: .Top = lblNote.Top + 45 + lblNote.Height
    End With
    With Me.LinTop
        .X1 = 0: .Y1 = Me.lblNote.Top + Me.lblNote.Height + 45 + Me.frmBaby.Height + mHeight
        .X2 = Me.ScaleWidth: .Y2 = .Y1
    End With
    With Me.rptList
        .Left = 150: .Width = lngScaleRight - .Left * 2
        .Top = Me.LinTop.Y1 + 90: .Height = lngScaleBottom - .Top - 180 - vfgThis.Height - lblNeaten.Height
    End With
    With LineBottom
        .X1 = 0: .Y1 = rptList.Top + rptList.Height + 45
        .X2 = Me.ScaleWidth: .Y2 = .Y1
    End With
    lblNeaten.Move 150, LineBottom.Y1 + 90
    vfgThis.Move 150, lblNeaten.Top + lblNeaten.Height + 50, lngScaleRight - 300
End Sub
 

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Edit_NewItem
        Control.Enabled = (Me.rptList.Rows.Count > 0)
        If Control.Enabled Then Control.Enabled = Not (Me.rptList.FocusedRow Is Nothing)
    Case conMenu_View_Show
        Control.Checked = rptList.PreviewMode
        Control.IconId = IIf(Control.Checked, 90002, 90001)
    Case conMenu_View_ShowAll
        Control.Visible = (InStr(";" & mstrPrivs & ";", ";ǿ����д;") > 0)
        If Control.Visible Then Control.Checked = mblnShowAll
        Control.IconId = IIf(Control.Checked, 90002, 90001)
    End Select
End Sub
Private Function Search()
 Dim searchKey As String, i As Long
 Dim rptRow As ReportRecord
 
    searchKey = txtSearchKey.Text
    For Each rptRow In Me.rptList.Records
        rptRow.Visible = True
        If InStr(CStr(UCase(rptRow(mc.����).Value)), UCase(searchKey)) > 0 Or InStr(CStr(UCase(rptRow(mc.����).Value)), UCase(searchKey)) > 0 Or searchKey = "" Then

        Else
            rptRow.Visible = False
        End If
    Next
    
    rptList.Populate
    If rptList.Rows.Count > 0 Then
        If rptList.Rows(0).GroupRow = False Then
            Set rptList.FocusedRow = rptList.Rows(0)
        ElseIf rptList.Rows(1).GroupRow = False Then
            Set rptList.FocusedRow = rptList.Rows(1)
        End If
    End If
End Function

Private Sub Form_Load()
    Dim cbrControl As CommandBarControl, cbrToolBar As CommandBar
    Dim rptCol As ReportColumn, lngCount As Long

    '-----------------------------------------------------
    '�ڲ��˵�����������
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    With cbsThis
        Set .Icons = zlCommFun.GetPubIcons
        .VisualTheme = xtpThemeOfficeXP
        With Me.cbsThis.Options
            .ShowExpandButtonAlways = False
            .ToolBarAccelTips = True
            .AlwaysShowFullMenus = False
            .IconsWithShadow = True '����VisualTheme����Ч
            .UseDisabledIcons = True
            .LargeIcons = False
            .SetIconSize False, 16, 16
            .UseSharedImageList = False 'ImageList��ʽʱ,��ͬһApp�й���,��AddImageList֮ǰ����ΪFalse
        End With
        .AddImageList img16
        .EnableCustomization False
        .ActiveMenuBar.Visible = False
    End With
    
    Set cbrToolBar = cbsThis.Add("������", xtpBarTop)
    cbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    cbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    cbrToolBar.ContextMenuPresent = False
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����(&N)"): cbrControl.STYLE = xtpButtonIconAndCaption
        Set cbrControl = .Add(xtpControlButton, conMenu_View_ShowAll, "��ʾ����"): cbrControl.STYLE = xtpButtonIconAndCaption: cbrControl.flags = xtpFlagRightAlign
        mblnShowAll = IIf(Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��ʾ����", "1")) = 1, True, False)
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Show, "��ʾ˵��"): cbrControl.STYLE = xtpButtonIconAndCaption: cbrControl.flags = xtpFlagRightAlign
    End With
    
    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type <> xtpControlLabel Then
            cbrControl.STYLE = xtpButtonIconAndCaption
        End If
    Next
    
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("N"), conMenu_Edit_NewItem
    End With
    
    '-----------------------------------------------------
    '�б���涨��
    '-----------------------------------------------------
    Me.BackColor = RGB(240, 240, 240)
    With Me.rptList
        Set rptCol = .Columns.Add(conColumn_ͼ��, "����", 18, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(conColumn_ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(conColumn_�¼�, "�¼�", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(conColumn_����, "����", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        
        .SetImageList Me.imgList
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        .PreviewMode = IIf(Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "�ļ�˵��", "1")) = 1, True, False)
        With .PaintManager
            .BackColor = Me.BackColor
            .NoItemsText = "û�п������ĵ��Ӳ�����¼..."
            .HighlightBackColor = RGB(135, 195, 255)
            .HighlightForeColor = RGB(0, 0, 0)
            .HorizontalGridStyle = xtpGridNoLines
            .SetPreviewIndent 18, 0, 8, 6
        End With
    End With
    
    With vfgThis
        .Clear
        .Cols = 4
        .TextMatrix(0, 0) = "ID"
        .TextMatrix(0, 1) = "����"
        .TextMatrix(0, 2) = "����ʱ��"
        .TextMatrix(0, 3) = "����"
        .ColWidth(0) = 0
        .ColWidth(1) = 1800
        .ColWidth(2) = 2200
        .ColWidth(3) = 400
        
        For lngCount = 0 To 2
            .ColAlignment(lngCount) = flexAlignLeftCenter
        Next
    End With
    mblnShowAll = False
End Sub
Private Sub rptList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim intNum As Integer
    On Error Resume Next
    With Me.rptList
        If .FocusedRow Is Nothing Then Exit Sub
        If .FocusedRow.Record(conColumn_ID) Is Nothing Then Exit Sub
        If CLng(.FocusedRow.Record(conColumn_ID).Value) = 0 Then Exit Sub
        If KeyCode <> vbKeyReturn Then Exit Sub
        If frmBaby.Visible Then
            intNum = Me.cboӤ��.ItemData(cboӤ��.ListIndex)
        End If
        RaiseEvent NewClick(CLng(.FocusedRow.Record(conColumn_ID).Value), intNum)
    End With
End Sub

Private Sub rptList_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim cbrPopupBar As CommandBar, cbrControl As CommandBarControl
    If Button <> vbRightButton Then Exit Sub
    Set cbrPopupBar = Me.cbsThis.Add("����", xtpBarPopup)
    With cbrPopupBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����(&N)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Show, "��ʾ�ļ�˵��(&H)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_ShowAll, "��ʾ�����ļ�(&A)")
    End With
    cbrPopupBar.ShowPopup
End Sub

Private Sub rptList_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
Dim intNum As Integer
    On Error Resume Next
    With Me.rptList
        If .FocusedRow Is Nothing Then Exit Sub
        If frmBaby.Visible Then
            intNum = Me.cboӤ��.ItemData(cboӤ��.ListIndex)
        End If
        RaiseEvent NewClick(.FocusedRow.Record(conColumn_ID).Value, intNum)
    End With
End Sub

'-----------------------------------------------------
'���幫������
'-----------------------------------------------------

Public Function zlOpenDefaultEPR(Optional ByVal bytKind As Byte = 1) As Boolean
    
    'bytKind=1���ﲡ��;=2���ﲡ��;=3���ﲡ��
    Dim intLoop As Integer
    Dim strTmp As String
    Dim strSQL As String
    Dim strSort As String
    Dim rsTmp As New ADODB.Recordset
    
    Select Case bytKind
    Case 1
        strTmp = "����"
        strSort = "Decode(b.�¼�,'����',1,'����',2,3),a.����,a.���"
    Case 2
        strTmp = "����"
        strSort = "Decode(b.�¼�,'����',1,'����',2,3),a.����,a.���"
    Case 3
        strTmp = "����"
        strSort = "Decode(b.�¼�,'����',1,'����',2,3),a.����,a.���"
    Case Else
        strTmp = "����"
        strSort = "Decode(b.�¼�,'����',1,2),a.����,a.���"
    End Select
    

    strSQL = "Select a.Id, a.����, a.���, a.����, a.˵��,b.�¼� " & _
            " From �����ļ��б� a,����ʱ��Ҫ�� b " & _
            " Where a.ID=b.�ļ�id(+) And Instr(';' || Zl_Out_Epr_Allowed([1], [2], [3],[4]) || ';', ';' || a.Id || ';') <> 0" & _
            " Order By " & strSort
            
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngPatiId, mlngVisit, mlngDeptId, mlngAdviceID)
        
    If rsTmp.BOF = False Then

        RaiseEvent NewClick(Val(rsTmp("ID").Value), 0)
        
        zlOpenDefaultEPR = True
        
    End If
        
End Function

Public Function zlRefList(ByVal bytScene As Byte, ByVal lngPatiID As Long, ByVal lngVisit As Long, ByVal lngDeptId As Long, _
                            Optional ByVal strPrivs As String, Optional ByVal lngAdviceID As Long) As Long
    '******************************************************************************************************************
    '���ܣ���ʾָ�����˿�����ָ���������ӵĵ��Ӳ�����¼�ļ��嵥
    '������ bytScene�����Ӳ�����¼���ӳ�����1-����,2-סԺ,3-����
    '       lngPatiId������id
    '       lngVisit�����˾����־�����ﲡ��Ϊ�Һ�ID��סԺ����Ϊ��ҳid
    '       lngDeptId������id
    '���أ������ӵĵ��Ӳ������ļ���Ŀ��Ϊ0��������
    '******************************************************************************************************************
    Dim strIDs As String, blnShowAll As Boolean, lngCount As Long
    Dim rsTemp As New ADODB.Recordset, rsLimit As New ADODB.Recordset
    Dim rptRcd As ReportRecord, rptItem As ReportRecordItem
    Dim rptRow As ReportRow
    Dim blnDisease As Boolean

    mbytScene = bytScene
    mstrPrivs = strPrivs
    mlngPatiId = lngPatiID
    mlngVisit = lngVisit
    mlngDeptId = lngDeptId
    mlngAdviceID = lngAdviceID
    
    blnDisease = (GetPrivFunc(glngSys, 1249) <> "")
    
    Err = 0: On Error GoTo errHand
    Select Case bytScene
    Case 1                  '����
        If mblnShowAll Then
            gstrSQL = "Select a.Id, a.����, a.���, a.����,zlspellcode(a.����) ����, a.˵��,b.�¼� " & _
                    " From �����ļ��б� a,����ʱ��Ҫ�� b,����Ӧ�ÿ��� c " & _
                    " Where a.ID=b.�ļ�id(+) And a.ID=c.�ļ�id(+) And a.���� In " & IIf(blnDisease, " (1,6) ", " (1,5,6) ") & _
                    " And A.����<>4 And (a.ͨ��=1 Or a.ͨ��=2 And c.����id=[2])" & _
                    " Order By a.����, a.���"
        Else
            gstrSQL = "Select Zl_Out_Epr_Allowed([1], [2], [3],[4]) IDS From Dual"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngPatiID, lngVisit, lngDeptId, mlngAdviceID)
            If rsTemp.EOF Then
                strIDs = ""
            Else
                strIDs = NVL(rsTemp!IDS)
            End If
            gstrSQL = "Select a.Id, a.����, a.���, a.����,zlspellcode(a.����) ����, a.˵��,b.�¼� " & _
                    " From �����ļ��б� a,����ʱ��Ҫ�� b " & _
                    " Where a.ID=b.�ļ�id(+) And A.����<>4 And Instr(';' || [1] || ';', ';' || a.Id || ';') <> 0 " & _
                     IIf(blnDisease, " And a.���� <> 5 ", "") & _
                    " Order By a.����, a.���"
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strIDs, lngDeptId)
    '------------------------------------------------------------------------------------------------------------------
        gstrSQL = "Select �ļ�ID ID,������� || '-' || �������� ��������, ����ʱ��, Decode(����,1,'��','��') ��д" & vbNewLine & _
                    "From ���Ӳ���ʱ��" & vbNewLine & _
                    "Where ����id = [1] And ��ҳid = [2] And ����id =[3] and �������� in " & IIf(blnDisease, " (1,6) ", " (1,5,6) ") & _
                    "And ������Դ = 1 And (Nvl(��ɼ�¼id, 0) = 0 And ���ʱ�� Is Null)" & vbNewLine & _
                    "Order By ����ʱ��"
        Set rsLimit = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngPatiID, lngVisit, lngDeptId, gstrUserName)
    Case 2                  'סԺ
        If mblnShowAll Then
        
            gstrSQL = "Select a.Id, a.����, a.���, a.����,zlspellcode(a.����) ����, a.˵��,b.�¼� " & _
                    " From �����ļ��б� a,����ʱ��Ҫ�� b,����Ӧ�ÿ��� c " & _
                    " Where a.ID=b.�ļ�id(+) And a.ID=c.�ļ�id(+) And a.���� In " & IIf(blnDisease, " (2,6) ", " (2,5,6) ") & _
                    " And A.����<>4 And (a.ͨ��=1 Or a.ͨ��=2 And c.����id=[2])" & _
                    " Order By a.����, a.���"
        Else
            gstrSQL = "Select Zl_In_Epr_Allowed([1], [2], [3]) IDS From Dual"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngPatiID, CLng(lngVisit), lngDeptId)
            If rsTemp.EOF Then
                strIDs = ""
            Else
                strIDs = NVL(rsTemp!IDS)
            End If
            gstrSQL = "Select Id, ����, ���, ����,zlspellcode(����) ����, ˵��,'' As �¼�" & _
                    " From �����ļ��б�" & _
                    " Where ����<>4 And Instr(';' || [1] || ';', ';' || Id || ';') <> 0" & _
                    IIf(blnDisease, " And ���� <> 5 ", "") & _
                    " Order By ����, ���"
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strIDs, lngDeptId)
    '------------------------------------------------------------------------------------------------------------------
        gstrSQL = "Select �ļ�ID ID,������� || '-' || �������� ��������, ����ʱ��, Decode(����,1,'��','��') ��д" & vbNewLine & _
                    "From ���Ӳ���ʱ��" & vbNewLine & _
                    "Where ����id = [1] And ��ҳid = [2] And ����id =[3] and �������� in " & IIf(blnDisease, " (2,6) ", " (2,5,6) ") & _
                    "And ������Դ = 2 And (Nvl(��ɼ�¼id, 0) = 0 And ���ʱ�� Is Null)" & vbNewLine & _
                    "Order By ����ʱ��"
        Set rsLimit = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngPatiID, lngVisit, lngDeptId, gstrUserName)
    Case 3                  '����
        If mblnShowAll Then
        
            gstrSQL = "Select a.Id, a.����, a.���, a.����,zlspellcode(a.����) ����, a.˵��,b.�¼� " & _
                    " From �����ļ��б� a,����ʱ��Ҫ�� b,����Ӧ�ÿ��� c " & _
                    " Where a.ID=b.�ļ�id(+) And a.ID=c.�ļ�id(+) And a.����=4 And (a.ͨ��=1 Or a.ͨ��=2 And c.����id=[2])" & _
                    " Order By a.����, a.���"
        Else
            gstrSQL = "Select Zl_Nurse_Epr_Allowed([1], [2], [3]) IDS From Dual"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngPatiID, CLng(lngVisit), lngDeptId)
            If rsTemp.EOF Then
                strIDs = ""
            Else
                strIDs = NVL(rsTemp!IDS)
            End If
            gstrSQL = "Select Id, ����, ���, ����,zlspellcode(����) ����, ˵��,'' As �¼�" & _
                    " From �����ļ��б�" & _
                    " Where ����<>4 And Instr(';' || [1] || ';', ';' || Id || ';') <> 0" & _
                    " Order By ����, ���"
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strIDs, lngDeptId)
    '------------------------------------------------------------------------------------------------------------------
        gstrSQL = "Select �ļ�ID ID,������� || '-' || �������� ��������, ����ʱ��, Decode(����,1,'��','��') ��д" & vbNewLine & _
                    "From ���Ӳ���ʱ��" & vbNewLine & _
                    "Where ����id = [1] And ��ҳid = [2] And ����id =[3] and ��������=4 And ������Դ = 2 And (Nvl(��ɼ�¼id, 0) = 0 And ���ʱ�� Is Null)" & vbNewLine & _
                    "Order By ����ʱ��"
        Set rsLimit = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngPatiID, lngVisit, lngDeptId, gstrUserName)
    End Select
    
    '------------------------------------------------------------------------------------------------------------------
    Me.rptList.Records.DeleteAll
    With rsTemp
        Do While Not .EOF
            Set rptRcd = Me.rptList.Records.Add()
            Select Case !����
            Case 1: Set rptItem = rptRcd.AddItem(CStr("1-���ﲡ��")): rptItem.Icon = !����
            Case 2: Set rptItem = rptRcd.AddItem(CStr("2-סԺ����")): rptItem.Icon = !����
            Case 3: Set rptItem = rptRcd.AddItem(CStr("3-�����¼")): rptItem.Icon = !����
            Case 4: Set rptItem = rptRcd.AddItem(CStr("4-������")): rptItem.Icon = !����
            Case 5: Set rptItem = rptRcd.AddItem(CStr("5-����֤������")): rptItem.Icon = !����
            Case 6: Set rptItem = rptRcd.AddItem(CStr("6-֪���ļ�")): rptItem.Icon = !����
            Case Else: rptRcd.AddItem ""
            End Select
            
            rptRcd.AddItem CStr(!ID)
            rptRcd.AddItem CStr("" & !�¼�)
            rptRcd.AddItem CStr(!��� & "-" & !����)
            rptRcd.AddItem CStr(!����)
            rptRcd.PreviewText = CStr("" & !˵��)
            .MoveNext
        Loop
    End With
    Me.rptList.SortOrder.Add rptList.Columns.Find(conColumn_����)
    Me.rptList.SortOrder(0).SortAscending = True
    Me.rptList.GroupsOrder.DeleteAll
    Me.rptList.GroupsOrder.Add Me.rptList.Columns.Find(0)
    Me.rptList.GroupsOrder(0).SortAscending = True
    Me.rptList.Populate
    For Each rptRow In Me.rptList.Rows
        If rptRow.GroupRow = True And rptRow.Index > 0 Then
           rptRow.Expanded = False
        End If
    Next
    If Me.rptList.Rows.Count > 0 Then
        Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
    End If
    
    
    With vfgThis
        .Clear
        .Cols = 4
        .TextMatrix(0, 0) = "ID"
        .TextMatrix(0, 1) = "����"
        .TextMatrix(0, 2) = "����ʱ��"
        .TextMatrix(0, 3) = "����"
        Set .DataSource = rsLimit
        .ColWidth(0) = 0
        .ColWidth(1) = 2000
        .ColWidth(2) = 1800
        .ColWidth(3) = 500
        
        For lngCount = 0 To 2
            .ColAlignment(lngCount) = flexAlignLeftCenter
        Next
    End With

    
    '------------------------------------------------------------------------------------------------------------------
    '����Ӥ���б�
    gstrSQL = "select ���,decode(Ӥ������,null,'Ӥ��'||���,Ӥ������)||' ����' ���� from ������������¼ where ����id = [1] And ��ҳid = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngPatiID, lngVisit)
    If rsTemp.RecordCount = 0 Then
        Me.cboӤ��.Clear
        Me.frmBaby.Visible = False
        mHeight = -Me.frmBaby.Height
    Else
        With rsTemp
            Me.cboӤ��.Clear
            Me.cboӤ��.AddItem ("���˲���")
            Do While Not .EOF
                Me.cboӤ��.AddItem (!����)
                Me.cboӤ��.ItemData(Me.cboӤ��.NewIndex) = NVL(!���, 0)
                .MoveNext
            Loop
        End With
        Me.cboӤ��.ListIndex = 0
        mHeight = 0
        Me.frmBaby.Visible = True
    End If
    zlRefList = rptList.Records.Count
    Call cbsThis_Resize
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    zlRefList = rptList.Records.Count
End Function
Private Sub txtSearchKey_Change()
    Call Search
End Sub

Private Sub txtSearchKey_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    Call Search
End Sub

Private Sub vfgThis_DblClick()
Dim intNum As Integer
    With Me.vfgThis
        If .TextMatrix(1, 0) = "" Then Exit Sub
        If frmBaby.Visible Then
            intNum = Me.cboӤ��.ItemData(cboӤ��.ListIndex)
        End If
        RaiseEvent NewClick(CLng(.TextMatrix(.Row, 0)), intNum)
    End With
End Sub


