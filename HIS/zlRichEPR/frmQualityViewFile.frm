VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Begin VB.Form frmQualityViewFile 
   Caption         =   "���Ҳ���������"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   10815
   Icon            =   "frmQualityViewFile.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   10815
   StartUpPosition =   3  '����ȱʡ
   Begin XtremeReportControl.ReportControl rptList 
      Height          =   3780
      Left            =   75
      TabIndex        =   0
      Top             =   660
      Width           =   3975
      _Version        =   589884
      _ExtentX        =   7011
      _ExtentY        =   6667
      _StockProps     =   0
      BorderStyle     =   2
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   6165
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmQualityViewFile.frx":6852
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13996
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   585
      Top             =   4455
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
            Picture         =   "frmQualityViewFile.frx":70E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQualityViewFile.frx":767E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQualityViewFile.frx":7C18
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQualityViewFile.frx":81B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQualityViewFile.frx":874C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQualityViewFile.frx":8CE6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vgdList 
      Height          =   900
      Left            =   150
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4995
      Visible         =   0   'False
      Width           =   1080
      _cx             =   1905
      _cy             =   1587
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   0   'False
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
      GridColor       =   -2147483632
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
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
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
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   1395
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgThis 
      Height          =   2175
      Left            =   4230
      TabIndex        =   3
      Top             =   675
      Width           =   3900
      _cx             =   6879
      _cy             =   3836
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
      GridColor       =   12632256
      GridColorFixed  =   12632256
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
      FormatString    =   $"frmQualityViewFile.frx":9280
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
   Begin VB.Image imgBG 
      Height          =   2295
      Left            =   8325
      Picture         =   "frmQualityViewFile.frx":9355
      Top             =   3690
      Visible         =   0   'False
      Width           =   2265
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   2160
      Top             =   180
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmQualityViewFile.frx":1A41F
      Left            =   900
      Top             =   225
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmQualityViewFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mCol
    ͼ�� = 0: ID: ����: ���: ����: ������
End Enum

Private Enum mCol2
    ID = 0: ����ID: ��ҳID: סԺ��: ����: ����: �Ա�: ����: סԺҽʦ: �ѱ�: ��Ժ����: ��Ժ����: ����: ������: ����ʱ��: ���ʱ��: ������: ����ʱ��: ���汾: ǩ������: �鵵��: �鵵����
End Enum

Private Enum mViewModeEnum
    ���в��� = 0: ������д����: ����ɲ���
End Enum

Private Enum Enum��������
    ���ﲡ�� = 1
    סԺ���� = 2
    ������ = 4
End Enum
Private mvar�������� As Enum��������

Const conPane_FileTab = 201
Const conPane_FileList = 202
Const conPane_Content = 203
Const conViewAll = 301
Const conViewInEditing = 302
Const conViewFinished = 303

'-----------------------------------------------------
'�������
'-----------------------------------------------------
Private mstrPrivs As String         '��ǰʹ����Ȩ�޴�
Private mstrKinds As String         '��ǰ������Ĳ������ʹ�

Private WithEvents mfrmContent As frmDockEPRContent
Attribute mfrmContent.VB_VarHelpID = -1

Private mlngCurFileId As Long       '��ǰ�ļ�ID
Private mlngDeptId As Long          '����ID
Private mstrDeptName As String      '��������
Private mstrFrom As String, mstrTo As String
Private mViewMode As mViewModeEnum  '��ͼģʽ   0-������д����    1-����ɲ���

'-----------------------------------------------------
'��ʱ����
'-----------------------------------------------------
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar

Dim rptCol As ReportColumn
Dim rptRcd As ReportRecord
Dim rptItem As ReportRecordItem
Dim rptRow As ReportRow

Dim rsTemp As New ADODB.Recordset
Dim lngCount As Long, lngCurRow As Long

Public Sub ShowMe(ByVal �������� As Long, _
    ByRef frmParent As Object, _
    ByVal lngDeptId As Long, _
    ByVal strDeptName As String, _
    ByVal strFrom As String, _
    ByVal strTo As String)
    '��ʾ������
    mvar�������� = ��������
    mlngDeptId = lngDeptId
    mstrDeptName = strDeptName
    mstrFrom = strFrom
    mstrTo = strTo
    Me.Caption = "���Ҳ��������� - [" & strDeptName & "]"
    Me.Show vbModeless, frmParent
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '-------------------------------------------------
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '       strSubhead����ӡ�ĸ�����
    '-------------------------------------------------
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    Set objPrint.Body = Me.vfgThis
    objPrint.Title.Text = Me.Caption
    objPrint.Title.Font.Name = "����"
    
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("����:" & mstrDeptName)
    Call objAppRow.Add("�ļ���:" & Me.rptList.FocusedRow.Record.Item(mCol.����).Value)
    Call objPrint.UnderAppRows.Add(objAppRow)
    
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("��ӡʱ��:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngFileID As Long
    On Error GoTo LL
    
    Select Case Control.ID
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview: Call zlRptPrint(0)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_File_Exit: Unload Me
    Case conMenu_View_ToolBar_Button
        Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each cbrControl In Me.cbsThis(2).Controls
            cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
        Me.cbsThis.RecalcLayout
    Case conMenu_View_StatusBar
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_Refresh
        Call zlRefList(mlngCurFileId)
    Case conViewAll
        mViewMode = ���в���
        Call rptList_SelectionChanged
    Case conViewInEditing
        mViewMode = ������д����
        Call rptList_SelectionChanged
    Case conViewFinished
        mViewMode = ����ɲ���
        Call rptList_SelectionChanged
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
    End Select
LL:
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel, conMenu_File_ExportToXML
        Control.Enabled = (Me.rptList.Records.Count <> 0)
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    Case conViewAll: Control.Checked = (mViewMode = ���в���)
    Case conViewInEditing: Control.Checked = (mViewMode = ������д����)
    Case conViewFinished: Control.Checked = (mViewMode = ����ɲ���)
    End Select
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    dkpMan.RecalcLayout
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_FileTab
        Item.Handle = Me.rptList.hWnd
    Case conPane_FileList
        Item.Handle = vfgThis.hWnd
    Case conPane_Content
        If mfrmContent Is Nothing Then Set mfrmContent = New frmEPRFileContent
        Item.Handle = mfrmContent.hWnd
    End Select
End Sub

Private Sub Form_Load()
    mViewMode = ���в���
    '-----------------------------------------------------
    'Ȩ�����ƴ����ƣ�����ͬʱ��������ģ�������gstrPrivs�仯�����¿�����Ч
    mstrPrivs = gstrPrivs
'    mstrKinds = ""
'    If InStr(1, mstrPrivs, "���ﲡ��") > 0 Then mstrKinds = mstrKinds & ",1"
'    If InStr(1, mstrPrivs, "סԺ����") > 0 Then mstrKinds = mstrKinds & ",2"
'    If InStr(1, mstrPrivs, "�����¼") > 0 Then mstrKinds = mstrKinds & ",3"
'    If InStr(1, mstrPrivs, "������") > 0 Then mstrKinds = mstrKinds & ",4"
'    If InStr(1, mstrPrivs, "����֤������") > 0 Then mstrKinds = mstrKinds & ",5"
'    If InStr(1, mstrPrivs, "֪���ļ�") > 0 Then mstrKinds = mstrKinds & ",6"
'    If mstrKinds <> "" Then mstrKinds = Mid(mstrKinds, 2)
    mstrKinds = "1,2,3,4,5,6"
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Me.cbsThis.Icons = frmPubIcons.imgPublic.Icons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    
    '-----------------------------------------------------
    '�˵�����
    Me.cbsThis.ActiveMenuBar.Title = "�˵�"
    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        Set cbrControl = .Add(xtpControlButton, conViewAll, "���в���"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conViewInEditing, "������д����")
        Set cbrControl = .Add(xtpControlButton, conViewFinished, "����ɲ���")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): cbrControl.BeginGroup = True
    End With
    
    '�����
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F6, conMenu_View_Jump
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '���ò����ò˵�
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
        .AddHiddenCommand conMenu_View_Jump
    End With
    '-----------------------------------------------------
    '����������
    Set cbrToolBar = Me.cbsThis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set cbrControl = .Add(xtpControlButton, conViewAll, "���в���"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conViewInEditing, "������д����")
        Set cbrControl = .Add(xtpControlButton, conViewFinished, "����ɲ���")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '-----------------------------------------------------
    '���ôʾ���ʾͣ������
    If mfrmContent Is Nothing Then Set mfrmContent = New frmDockEPRContent
    
    Dim panFileTab As Pane, panModel As Pane, panCompend As Pane
    Set panFileTab = dkpMan.CreatePane(conPane_FileTab, 180, 400, DockLeftOf, Nothing)
    panFileTab.Title = "�����ļ��б�"
    panFileTab.Options = PaneNoCaption
    
    Set panModel = dkpMan.CreatePane(conPane_FileList, 400, 200, DockRightOf, Nothing)
    panModel.Title = "�����嵥"
    panModel.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable

    Set panCompend = dkpMan.CreatePane(conPane_Content, 400, 300, DockBottomOf, panModel)
    panCompend.Title = "��������"
    panCompend.Options = PaneNoCaption
    
    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
    
    '-----------------------------------------------------
    With Me.rptList
        Set rptCol = .Columns.Add(mCol.ͼ��, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mCol.ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.����, "����", 90, False): rptCol.Editable = False: rptCol.Groupable = True: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.���, "���", 50, False): rptCol.Editable = False: rptCol.Groupable = False: .SortOrder.Add rptCol
        Set rptCol = .Columns.Add(mCol.����, "����", 150, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.������, "������д/�����", 100, True): rptCol.Editable = False: rptCol.Groupable = False
        
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
    End With
    
    '-----------------------------------------------------
    '����ָ�
    Call RestoreWinState(Me, App.ProductName)
    '-----------------------------------------------------
    '����װ��
    If mstrKinds = "" Then
        DoEvents
        Me.stbThis.Panels(2).Text = "�㲻�߱������ļ��������Ȩ��"
    Else
        lngCount = Me.zlRefList()
        Me.stbThis.Panels(2).Text = "����" & lngCount & "�ݲ����ļ�"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload mfrmContent
    Set mfrmContent = Nothing
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub mfrmContent_DblClick()
    Dim f As New frmEPRView
    Dim lngFileID As Long
    lngFileID = Val(vfgThis.TextMatrix(vfgThis.Row, mCol2.ID))
    If lngFileID > 0 Then f.ShowMe Me, lngFileID
End Sub

Private Sub rptList_SelectionChanged()
    With Me.rptList
        If .FocusedRow Is Nothing Then
            mlngCurFileId = 0
        ElseIf .FocusedRow.GroupRow = True Then
            mlngCurFileId = 0
        Else
            mlngCurFileId = .FocusedRow.Record.Item(mCol.ID).Value  '��ȡ��ǰ�ļ�ID
        End If
    End With
    FillGrid mstrFrom, mstrTo
End Sub

Private Sub vfgThis_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    dkpMan.RecalcLayout
End Sub

Public Function zlRefList(Optional lngFileID As Long) As Long
    '���ܣ�ˢ��װ��ָ������Ĳ����ļ��嵥������λ��ָ�����ļ���
    Dim strGroups As String
    
    gstrSQL = "SELECT L.ID, L.����, L.���, L.����, P.������д, P.����� " & _
        "FROM �����ļ��б� L, " & _
        "     (SELECT M.�ļ�id, SUM(Decode(M.���ʱ��, NULL, 1, 0)) AS ������д, " & _
        "              SUM(Decode(M.���ʱ��, NULL, Decode(Sign(SYSDATE - M.����ʱ�� - 1), 1, 1, 0), 0)) AS ��д��ʱ, " & _
        "              SUM(Decode(M.���ʱ��, NULL, 0, 1)) AS �����, " & _
        "              SUM(Decode(M.���ʱ��, Null, 0, Decode(NVL(M.ǩ������, 0), 0, 1, 0))) As �����޶� " & _
        "       FROM ���Ӳ�����¼ M, ������ҳ N " & _
        "       WHERE M.����id = N.����id AND Nvl(N.��ҳid, 0) <> 0 AND Nvl(N.״̬, 0) <> 1 AND " & _
        "             N.��Ժ���� BETWEEN [3] And [4] AND M.����id = [1] AND M.�������� = [2] AND " & _
        "             M.��ҳID = N.��ҳID " & _
        "       GROUP BY �ļ�id) P " & _
        "WHERE L.���� = [2] AND P.�ļ�id(+) = L.ID"
            
    Err = 0: On Error GoTo errHand
    Dim lngNum1 As Long, lngNum2 As Long
    Set rsTemp = OpenSQLRecord(gstrSQL, Me.Caption, mlngDeptId, mvar��������, CDate(Format(mstrFrom, "YYYY-MM-DD")), CDate(Format(mstrTo, "YYYY-MM-DD") & " 23:59:59"))
    
    Me.rptList.Records.DeleteAll
    With rsTemp
        strGroups = ""
        Do While Not .EOF
            If InStr(1, strGroups, !����) = 0 Then strGroups = strGroups & "," & !����
            Set rptRcd = Me.rptList.Records.Add()
            Set rptItem = rptRcd.AddItem(CStr(!����)): rptItem.Icon = rptItem.Value - 1
            rptRcd.AddItem CStr(!ID)
            Select Case !����
            Case 1: rptRcd.AddItem CStr("1-���ﲡ��")
            Case 2: rptRcd.AddItem CStr("2-סԺ����")
            Case 3: rptRcd.AddItem CStr("3-�����¼")
            Case 4: rptRcd.AddItem CStr("4-������")
            Case 5: rptRcd.AddItem CStr("5-����֤������")
            Case 6: rptRcd.AddItem CStr("6-֪���ļ�")
            Case Else: rptRcd.AddItem ""
            End Select
            rptRcd.AddItem Val(CStr(!���))
            rptRcd.AddItem CStr(!����)
            lngNum1 = NVL(!������д, 0)
            lngNum2 = NVL(!�����, 0)
            rptRcd.AddItem IIf(lngNum1 = 0 And lngNum2 = 0, "", lngNum1 & "/" & lngNum2)
            .MoveNext
        Loop
        If strGroups <> "" Then strGroups = Mid(strGroups, 2)
    End With
    With Me.rptList
        If UBound(Split(strGroups, ",")) < 1 Then
            .GroupsOrder.DeleteAll
        ElseIf .GroupsOrder.Count = 0 Then
            .GroupsOrder.Add .Columns.Find(mCol.����)
            .GroupsOrder(0).SortAscending = True
        End If
        .Populate
    End With
    
    If lngFileID <> 0 Then
        For Each rptRow In Me.rptList.Rows
            If rptRow.GroupRow = False Then
                If Val(rptRow.Record(mCol.ID).Value) = lngFileID Then
                    Set Me.rptList.FocusedRow = rptRow: Exit For
                End If
            End If
        Next
    End If
    If Me.rptList.Rows.Count > 0 Then
        If Me.rptList.FocusedRow Is Nothing Then Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
        If Me.rptList.FocusedRow.GroupRow Then
            lngFileID = 0
        Else
            lngFileID = Me.rptList.FocusedRow.Record.Item(mCol.ID).Value
        End If
    Else
        lngFileID = 0
    End If
    
    zlRefList = Me.rptList.Records.Count
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefList = Me.rptList.Records.Count
    lngFileID = 0
End Function

Private Sub InitGrid()
    Dim i As Long
    With Me.vfgThis
        .Clear
        .Rows = 1
        .FixedRows = 1
        .Cols = 22
        .RowHeightMin = 300
        .WallPaper = imgBG.Picture
        .WallPaperAlignment = flexPicAlignRightBottom
        
'        .BackColorAlternate = RGB(240, 240, 255)
        .BackColorSel = RGB(125, 125, 255)
        .ForeColorSel = vbWhite
        .Sort = flexSortCustom
        
        .TextMatrix(0, mCol2.ID) = "ID"
        .TextMatrix(0, mCol2.����ID) = "����ID"
        .TextMatrix(0, mCol2.��ҳID) = "��ҳID"
        .TextMatrix(0, mCol2.סԺ��) = "סԺ��"
        .TextMatrix(0, mCol2.����) = "����"
        .TextMatrix(0, mCol2.����) = "����"
        .TextMatrix(0, mCol2.�Ա�) = "�Ա�"
        .TextMatrix(0, mCol2.����) = "����"
        .TextMatrix(0, mCol2.סԺҽʦ) = "סԺҽʦ"
        .TextMatrix(0, mCol2.�ѱ�) = "�ѱ�"
        .TextMatrix(0, mCol2.��Ժ����) = "��Ժ����"
        .TextMatrix(0, mCol2.��Ժ����) = "��Ժ����"
        .TextMatrix(0, mCol2.����) = "����"
        .TextMatrix(0, mCol2.������) = "������"
        .TextMatrix(0, mCol2.����ʱ��) = "����ʱ��"
        .TextMatrix(0, mCol2.���ʱ��) = "���ʱ��"
        .TextMatrix(0, mCol2.������) = "������"
        .TextMatrix(0, mCol2.����ʱ��) = "����ʱ��"
        .TextMatrix(0, mCol2.���汾) = "���汾"
        .TextMatrix(0, mCol2.ǩ������) = "ǩ������"
        .TextMatrix(0, mCol2.�鵵��) = "�鵵��"
        .TextMatrix(0, mCol2.�鵵����) = "�鵵����"
        
'        .MergeCol(mCol2.����ID) = True
'        .MergeCol(mCol2.��ҳID) = True
'        .MergeCol(mCol2.סԺ��) = True
'        .MergeCol(mCol2.����) = True
'        .MergeCol(mCol2.����) = True
'        .MergeCol(mCol2.�Ա�) = True
'        .MergeCol(mCol2.����) = True
'        .MergeCol(mCol2.סԺҽʦ) = True
'        .MergeCol(mCol2.�ѱ�) = True
'        .MergeCol(mCol2.��Ժ����) = True
'        .MergeCol(mCol2.��Ժ����) = True
'        .MergeCol(mCol2.����) = True
'
'        .MergeCells = flexMergeRestrictColumns
        
        For i = 0 To 21
            .Cell(flexcpAlignment, 0, i) = flexAlignCenterCenter
        Next
        .ColWidth(mCol2.ID) = 0
        .ColWidth(mCol2.����ID) = 800
        .ColWidth(mCol2.��ҳID) = 600
        .ColWidth(mCol2.סԺ��) = 1100
        .ColWidth(mCol2.����) = 800
        .ColWidth(mCol2.����) = 800
        .ColWidth(mCol2.�Ա�) = 600
        .ColWidth(mCol2.����) = 600
        .ColWidth(mCol2.סԺҽʦ) = 800
        .ColWidth(mCol2.�ѱ�) = 800
        .ColWidth(mCol2.��Ժ����) = 1600
        .ColWidth(mCol2.��Ժ����) = 1600
        .ColWidth(mCol2.����) = 600
        .ColWidth(mCol2.������) = 800
        .ColWidth(mCol2.����ʱ��) = 1600
        .ColWidth(mCol2.���ʱ��) = 1600
        .ColWidth(mCol2.������) = 800
        .ColWidth(mCol2.����ʱ��) = 1600
        .ColWidth(mCol2.���汾) = 800
        .ColWidth(mCol2.ǩ������) = 800
        .ColWidth(mCol2.�鵵��) = 800
        .ColWidth(mCol2.�鵵����) = 1600
    End With
End Sub

Private Sub FillGrid(ByVal strFrom As String, ByVal strTo As String)
    With mfrmContent.edtThis
        .ForceEdit = True
        .ReadOnly = False
        .NewDoc
        .ReadOnly = True
        .ForceEdit = False
    End With
    '�������
    Dim Rs As ADODB.Recordset, i As Long, lngCount(1 To 10) As Long
    gstrSQL = "select l.ID, c.����ID, c.��ҳID, c.סԺ��, c.����, c.����, c.�Ա�, c.����, " & _
        "       c.סԺҽʦ, c.�ѱ�, c.��Ժ����, c.��Ժ����, c.����, l.������, " & _
        "       l.����ʱ��, l.���ʱ��, l.������, l.����ʱ��, l.���汾, l.ǩ������, " & _
        "       l.�鵵�� , l.�鵵����, l.����״̬ " & _
        "  from ���Ӳ�����¼ l, ������Ϣ b, " & _
        "       (Select A.����ID, B.��ҳID, A.סԺ��, A.����, A.�Ա�, A.����, " & _
        "                B.סԺҽʦ, B.��Ժ���� as ����, B.�ѱ�, B.��Ժ����, B.��Ժ����, " & _
        "                b.״̬ , b.���� " & _
        "           From ������Ϣ A, ������ҳ B " & _
        "          Where a.����ID = b.����ID " & _
        "            And Nvl(B.��ҳID, 0) <> 0 " & _
        "            And Nvl(B.״̬, 0) <> 1 " & _
        "            And " & _
        "                (B.��Ժ���� Between [1] And [2]) " & _
        "          Order by סԺ�� Desc, ��ҳID Desc) c " & _
        " Where c.����ID = B.����ID " & _
        "   and l.����id = B.����id and l.��ҳID = c.��ҳID " & _
        "   and l.����id = " & mlngDeptId & " and l.�ļ�id = " & mlngCurFileId & _
        IIf(mViewMode = ���в���, "", IIf(mViewMode = ������д����, " and l.���ʱ�� is null ", " and l.���ʱ�� is not null ")) & _
        " order by ����ID,��ҳID,����ʱ�� "
    
    Set Rs = OpenSQLRecord(gstrSQL, Me.Caption, CDate(Format(strFrom, "YYYY-MM-DD")), CDate(Format(strTo, "YYYY-MM-DD") & " 23:59:59"))
    Call InitGrid
    Me.vfgThis.Rows = 1 + Rs.RecordCount
    
    Me.stbThis.Panels(2).Text = IIf(mViewMode = ������д����, "������д", IIf(mViewMode = ���в���, "���в���", "�Ѿ���д")) & Rs.RecordCount & "�ݲ���"
    i = 1
    Do While Not Rs.EOF
        With Me.vfgThis
            .TextMatrix(i, mCol2.ID) = NVL(Rs("ID"), 0)
            .TextMatrix(i, mCol2.����ID) = NVL(Rs("����ID"), 0)
            .TextMatrix(i, mCol2.��ҳID) = NVL(Rs("��ҳID"), 0)
            .TextMatrix(i, mCol2.סԺ��) = NVL(Rs("סԺ��"), 0)
            .TextMatrix(i, mCol2.����) = NVL(Rs("����"), 0)
            .TextMatrix(i, mCol2.����) = NVL(Rs("����"))
            .TextMatrix(i, mCol2.�Ա�) = NVL(Rs("�Ա�"))
            .TextMatrix(i, mCol2.����) = NVL(Rs("����"))
            .TextMatrix(i, mCol2.סԺҽʦ) = NVL(Rs("סԺҽʦ"))
            .TextMatrix(i, mCol2.�ѱ�) = NVL(Rs("�ѱ�"))
            .TextMatrix(i, mCol2.��Ժ����) = Format(NVL(Rs("��Ժ����")), "yyyy-MM-DD HH:nn")
            .TextMatrix(i, mCol2.��Ժ����) = Format(NVL(Rs("��Ժ����")), "yyyy-MM-DD HH:nn")
            .TextMatrix(i, mCol2.����) = NVL(Rs("����"))
            .TextMatrix(i, mCol2.������) = NVL(Rs("������"))
            .TextMatrix(i, mCol2.����ʱ��) = Format(NVL(Rs("����ʱ��")), "yyyy-MM-DD HH:nn")
            .TextMatrix(i, mCol2.���ʱ��) = Format(NVL(Rs("���ʱ��")), "yyyy-MM-DD HH:nn")
            .TextMatrix(i, mCol2.������) = NVL(Rs("������"))
            .TextMatrix(i, mCol2.����ʱ��) = Format(NVL(Rs("����ʱ��")), "yyyy-MM-DD HH:nn")
            .TextMatrix(i, mCol2.���汾) = NVL(Rs("���汾"))
            .TextMatrix(i, mCol2.ǩ������) = NVL(Rs("ǩ������"))
            .TextMatrix(i, mCol2.�鵵��) = NVL(Rs("�鵵��"))
            .TextMatrix(i, mCol2.�鵵����) = Format(NVL(Rs("�鵵����")), "yyyy-MM-DD HH:nn")
        End With
        Rs.MoveNext
        i = i + 1
    Loop
    Rs.Close
    Set Rs = Nothing
    If vfgThis.Rows > 1 Then vfgThis.Row = 1: Call vfgThis_RowColChange
End Sub

Private Sub vfgThis_RowColChange()
    Dim lngFileID As Long
    lngFileID = Val(vfgThis.TextMatrix(vfgThis.Row, mCol2.ID))
    If lngFileID > 0 Then mfrmContent.zlRefresh lngFileID
End Sub

