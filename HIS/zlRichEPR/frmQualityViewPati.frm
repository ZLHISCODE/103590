VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmQualityViewPati 
   Caption         =   "����Ժ���˲���������"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   10875
   Icon            =   "frmQualityViewPati.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   10875
   StartUpPosition =   3  '����ȱʡ
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6555
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmQualityViewPati.frx":6852
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14102
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
            Picture         =   "frmQualityViewPati.frx":70E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQualityViewPati.frx":767E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQualityViewPati.frx":7C18
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQualityViewPati.frx":81B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQualityViewPati.frx":874C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQualityViewPati.frx":8CE6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vgdList 
      Height          =   900
      Left            =   150
      TabIndex        =   1
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
      Left            =   3240
      TabIndex        =   2
      Top             =   630
      Width           =   2820
      _cx             =   4974
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
      FormatString    =   $"frmQualityViewPati.frx":9280
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
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   2175
      Left            =   225
      TabIndex        =   3
      Top             =   630
      Width           =   2820
      _cx             =   4974
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
      FormatString    =   $"frmQualityViewPati.frx":9355
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   2655
      Top             =   180
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmQualityViewPati.frx":942A
      Left            =   900
      Top             =   225
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
   Begin VB.Image imgBG 
      Height          =   2295
      Left            =   8460
      Picture         =   "frmQualityViewPati.frx":943E
      Top             =   4095
      Visible         =   0   'False
      Width           =   2265
   End
End
Attribute VB_Name = "frmQualityViewPati"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mColPaiList
    ����ID = 0: סԺ��: ����: �Ա�: ����: סԺҽʦ: ����: �ѱ�: ��Ժ����: ��Ժ����
End Enum

Private Enum mColDetail
    ��־ = 0: �¼���ʱ��: Ӧд����: ����: Ҫ��ʱ��: ���ʱ��: ��ɼ�¼id: ��ǰʱ��: ������: ��ע˵��
End Enum

Private Enum Enum��������
    ���ﲡ�� = 1
    סԺ���� = 2
    ������ = 4
End Enum
Private mvar�������� As Enum��������

'-----------------------------------------------------
'�������
'-----------------------------------------------------
Private mstrPrivs As String         '��ǰʹ����Ȩ�޴�
Private mstrKinds As String         '��ǰ������Ĳ������ʹ�

Const conPane_PatiList = 201
Const conPane_Detail = 202
Const conPane_Content = 203

Private WithEvents mfrmContent As frmDockEPRContent
Attribute mfrmContent.VB_VarHelpID = -1

Private mlngCurFileId As Long       '��ǰ�ļ�ID
Private mlngDeptId As Long          '����ID
Private mstrDeptName As String      '��������
Private mstrFrom As String, mstrTo As String

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
    Me.Caption = "����Ժ���˲��������� - [" & strDeptName & "]"
    Call FillPatiList(mstrFrom, mstrTo)
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
        Control.Enabled = (Me.vfgThis.Records.Count <> 0)
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    End Select
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    dkpMan.RecalcLayout
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_PatiList
        Item.Handle = vfgList.hWnd
    Case conPane_Detail
        Item.Handle = vfgThis.hWnd
    Case conPane_Content
        If mfrmContent Is Nothing Then Set mfrmContent = New frmEPRFileContent
        Item.Handle = mfrmContent.hWnd
    End Select
End Sub

Private Sub Form_Load()
    '-----------------------------------------------------
    'Ȩ�����ƴ����ƣ�����ͬʱ��������ģ�������gstrPrivs�仯�����¿�����Ч
    mstrPrivs = gstrPrivs
    mstrKinds = ""
    If InStr(1, mstrPrivs, "���ﲡ��") > 0 Then mstrKinds = mstrKinds & ",1"
    If InStr(1, mstrPrivs, "סԺ����") > 0 Then mstrKinds = mstrKinds & ",2"
    If InStr(1, mstrPrivs, "�����¼") > 0 Then mstrKinds = mstrKinds & ",3"
    If InStr(1, mstrPrivs, "������") > 0 Then mstrKinds = mstrKinds & ",4"
    If InStr(1, mstrPrivs, "����֤������") > 0 Then mstrKinds = mstrKinds & ",5"
    If InStr(1, mstrPrivs, "֪���ļ�") > 0 Then mstrKinds = mstrKinds & ",6"
    If mstrKinds <> "" Then mstrKinds = Mid(mstrKinds, 2)
'    mstrKinds = "1,2,3,4,5,6"
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '-----------------------------------------------------
    '���ôʾ���ʾͣ������
    If mfrmContent Is Nothing Then Set mfrmContent = New frmDockEPRContent
    
    Dim panList As Pane, panCompend As Pane, lngCount As Long
    Set panList = dkpMan.CreatePane(conPane_PatiList, 160, 400, DockLeftOf, Nothing)
    panList.Title = "����Ժ�����б�"
    panList.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable

    Set panList = dkpMan.CreatePane(conPane_Detail, 400, 200, DockRightOf, Nothing)
    panList.Title = "��ϸ���"
    panList.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable

    Set panCompend = dkpMan.CreatePane(conPane_Content, 400, 300, DockBottomOf, panList)
    panCompend.Title = "��������"
    panCompend.Options = PaneNoCaption
    
    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
        
    '-----------------------------------------------------
    '����ָ�
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload mfrmContent
    Set mfrmContent = Nothing
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub mfrmContent_DblClick()
    Dim f As New frmEPRView
    Dim lngFileID As Long
    lngFileID = Val(vfgThis.TextMatrix(vfgThis.Row, mColDetail.��ɼ�¼id))
    If lngFileID > 0 Then f.ShowMe Me, lngFileID
End Sub

Private Sub vfgList_Click()
    Call vfgList_RowColChange
End Sub

Private Sub vfgList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    dkpMan.RecalcLayout
End Sub

Private Sub vfgList_RowColChange()
    Dim lngPatiId As Long
    lngPatiId = Val(vfgList.TextMatrix(vfgList.Row, mColPaiList.����ID))
    If lngPatiId > 0 Then FillDetail lngPatiId, 1, mvar��������
End Sub

Private Sub vfgThis_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    dkpMan.RecalcLayout
End Sub

Private Sub FillPatiList(ByVal strFrom As String, ByVal strTo As String)
    '��䲡���б�
    Select Case mvar��������
    Case ���ﲡ��
        'Ӧ�ôӹҺż�¼����ȡ��Ϣ
        gstrSQL = "SELECT A.����id, A.�����, A.����, A.�Ա�, A.����, B.����, B.ִ����, " & _
            "       To_Char(B.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi ') AS �Ǽ�ʱ�� " & _
            "FROM ������Ϣ A, ���˹Һż�¼ B " & _
            "WHERE A.����id = B.����id AND A.����� = B.����� AND Nvl(B.ִ��״̬, 0) <> 0 AND " & _
            "      (A.����ʱ�� BETWEEN [2] AND [3]" & _
            " )" & _
            "ORDER BY A.����� DESC"
    Case סԺ����
        gstrSQL = "Select A.����ID, A.סԺ��, A.����, A.�Ա�, A.����, B.סԺҽʦ, " & _
            "       B.��Ժ���� as ����, B.�ѱ�, To_Char(B.��Ժ����, 'yyyy-mm-dd hh24:mi ') As ��Ժ���� , To_Char(B.��Ժ����, 'yyyy-mm-dd hh24:mi ') As ��Ժ���� " & _
            " From ������Ϣ A, ������ҳ B " & _
            " Where a.����ID = b.����ID " & _
            "   and A.��ǰ����ID = [1] " & _
            "   And Nvl(B.��ҳID, 0) =1 " & _
            "   And Nvl(B.״̬, 0) <> 1 " & _
            "   And (B.��Ժ���� Between [2] And [3]) " & _
            " Order by B.סԺҽʦ Desc, סԺ�� Desc, ��ҳID Desc"
    Case ������
        gstrSQL = "Select A.����ID, A.סԺ��, A.����, A.�Ա�, A.����, B.סԺҽʦ, " & _
            "       B.��Ժ���� as ����, B.�ѱ�, To_Char(B.��Ժ����, 'yyyy-mm-dd hh24:mi ') As ��Ժ���� , To_Char(B.��Ժ����, 'yyyy-mm-dd hh24:mi ') As ��Ժ���� " & _
            " From ������Ϣ A, ������ҳ B " & _
            " Where a.����ID = b.����ID " & _
            "   and A.��ǰ����ID = [1] " & _
            "   And Nvl(B.��ҳID, 0) =1 " & _
            "   And Nvl(B.״̬, 0) <> 1 " & _
            "   And (B.��Ժ���� Between [2] And [3]) " & _
            " Order by B.סԺҽʦ Desc, סԺ�� Desc, ��ҳID Desc"
    End Select
    Dim strSQL As String
    Dim i As Long, j As Long
    Dim lngCount(1 To 6) As Long, strState As String    '������ʾ���˷���ͳ����Ŀ

    Set rsTemp = OpenSQLRecord(gstrSQL, Me.Caption, mlngDeptId, CDate(Format(mstrFrom, "YYYY-MM-DD")), CDate(Format(mstrTo, "YYYY-MM-DD") & " 23:59:59"))
    With Me.vfgList
        .Clear
        Set .DataSource = rsTemp
        .ColWidth(mColDetail.��־) = 0
    End With
    If vfgList.Rows > 1 Then vfgList.Row = 1: Call vfgList_RowColChange
End Sub

Private Sub FillDetail(ByVal lngPatiId As Long, ByVal lngPageId As Long, ByVal intKind As Integer)
    '��䲡��ʱ�޼����ϸ
    With mfrmContent.edtThis
        .ForceEdit = True
        .ReadOnly = False
        .NewDoc
        .ReadOnly = True
        .ForceEdit = False
    End With
    '---------------------------------------------------
    
    'ִ��ʱ�޼������
    gstrSQL = "Zl_����ʱ�޼��_Neaten(" & lngPatiId & "," & lngPageId & "," & intKind & ")"
    Call SQLTest(App.ProductName, Me.Caption, gstrSQL): gcnOracle.Execute gstrSQL, , adCmdStoredProc: Call SQLTest
    
    '1-���ﲡ��;2-סԺ����;3-�����¼;4-������;
    Dim lngCount As Long, lngBalance As Long, lngDay As Long, lngHour As Long
    gstrSQL = "Select To_Char(B.�¼�ʱ��, 'yyyy-mm-dd hh24:mi ') || B.�䶯�¼� As �¼���ʱ��, " & _
        "       B.������� || '-' || B.�������� As Ӧд����, " & _
        "       Decode(B.Ψһ, 1, '��д', '��' || B.���ں� || '����д') As ����, " & _
        "       B.Ҫ��ʱ��, B.���ʱ��, B.��ɼ�¼id, Sysdate As ��ǰʱ��, B.������, " & _
        "       '' As ��ע˵�� " & _
        " From ������Ϣ A, ����ʱ�޼�� B " & _
        " Where a.����ID = b.����ID " & _
        "   and A.����ID = [1] And B.��ҳID=[2] " & _
        "   And (B.�������� = [3] Or B.�������� in (5, 6) And [3] <> 4) " & _
        "   And B.Ҫ��ʱ�� - Sysdate < 2 " & _
        " Order By B.����ID, B.��ҳID, B.��������, B.�¼�ʱ��"
    Set rsTemp = OpenSQLRecord(gstrSQL, Me.Caption, lngPatiId, lngPageId, intKind)
    With Me.vfgThis
        .Clear
        .FixedCols = 1
        Set .DataSource = rsTemp
        '��־ = 0: �¼���ʱ��: Ӧд����: ����: Ҫ��ʱ��: ���ʱ��: ��ɼ�¼ID: ��ǰʱ��: ������: ��ע˵��
        
        .MergeCells = flexMergeFree: .MergeCol(mColDetail.�¼���ʱ��) = True: .MergeCol(mColDetail.Ӧд����) = True
        .ColWidth(mColDetail.��־) = 300: .ColWidth(mColDetail.�¼���ʱ��) = 2800: .ColWidth(mColDetail.Ӧд����) = 2000
        .ColWidth(mColDetail.����) = 1100: .ColWidth(mColDetail.Ҫ��ʱ��) = 1100
        .ColWidth(mColDetail.���ʱ��) = 0: .ColWidth(mColDetail.��ɼ�¼id) = 0: .ColWidth(mColDetail.��ǰʱ��) = 0
        .ColWidth(mColDetail.������) = 900: .ColWidth(mColDetail.��ע˵��) = 2200
        
        .FixedAlignment(mColDetail.��־) = flexAlignCenterCenter
        For lngCount = .FixedCols To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            .ColAlignment(lngCount) = flexAlignLeftTop
        Next
        For lngCount = .FixedRows To .Rows - 1
            If .TextMatrix(lngCount, mColDetail.���ʱ��) = "" Then
                If .TextMatrix(lngCount, mColDetail.��ɼ�¼id) = "" Then
                    .TextMatrix(lngCount, mColDetail.��ע˵��) = "δ��д"
                Else
                    .TextMatrix(lngCount, mColDetail.��ע˵��) = "������д"
                End If
                lngBalance = Int((CDate(.TextMatrix(lngCount, mColDetail.��ǰʱ��)) - CDate(.TextMatrix(lngCount, mColDetail.Ҫ��ʱ��))) * 24)
                .TextMatrix(lngCount, mColDetail.��־) = "��"
                If lngBalance >= 0 Then
                    .Cell(flexcpForeColor, lngCount, mColDetail.��־, lngCount, mColDetail.��־) = RGB(255, 0, 0)
                    If lngBalance > 24 Then
                        '����24Сʱ��������
                        lngDay = lngBalance / 24
                        lngHour = lngBalance Mod 24
                        .TextMatrix(lngCount, mColDetail.��ע˵��) = .TextMatrix(lngCount, mColDetail.��ע˵��) & IIf(lngBalance = 0, "", ",�ѳ���" & lngDay & "��" & lngHour & "Сʱ")
                    Else
                        .TextMatrix(lngCount, mColDetail.��ע˵��) = .TextMatrix(lngCount, mColDetail.��ע˵��) & IIf(lngBalance = 0, "", ",�ѳ���" & lngBalance & "Сʱ")
                    End If
                    .Cell(flexcpForeColor, lngCount, mColDetail.��ע˵��, lngCount, mColDetail.��ע˵��) = RGB(255, 0, 0)
                Else
                    If Abs(lngBalance) < 4 Then
                        .Cell(flexcpForeColor, lngCount, mColDetail.��־, lngCount, mColDetail.��־) = RGB(128, 128, 0)
                        .TextMatrix(lngCount, mColDetail.��ע˵��) = .TextMatrix(lngCount, mColDetail.��ע˵��) & ",ʣ��" & Abs(lngBalance) & "Сʱ,�뾡�����"
                    Else
                        .Cell(flexcpForeColor, lngCount, mColDetail.��־, lngCount, mColDetail.��־) = RGB(0, 0, 255)
                        .TextMatrix(lngCount, mColDetail.��ע˵��) = .TextMatrix(lngCount, mColDetail.��ע˵��) & ",ʣ��" & Abs(lngBalance) & "Сʱ,�밴ʱ���"
                    End If
                End If
            Else
                lngBalance = Int((CDate(.TextMatrix(lngCount, mColDetail.���ʱ��)) - CDate(.TextMatrix(lngCount, mColDetail.Ҫ��ʱ��))) * 24)
                If lngBalance > 0 Then
                    .TextMatrix(lngCount, mColDetail.��־) = "��"
                    .Cell(flexcpForeColor, lngCount, mColDetail.��־, lngCount, mColDetail.��־) = RGB(255, 0, 0)
                    .TextMatrix(lngCount, mColDetail.��ע˵��) = "���,������" & lngBalance & "Сʱ"
                    .Cell(flexcpForeColor, lngCount, mColDetail.��ע˵��, lngCount, mColDetail.��ע˵��) = RGB(255, 0, 0)
                Else
                    .TextMatrix(lngCount, mColDetail.��ע˵��) = "�������"
                End If
            End If
            .TextMatrix(lngCount, mColDetail.Ҫ��ʱ��) = Format(.TextMatrix(lngCount, mColDetail.Ҫ��ʱ��), "MM-dd hh:mm")
            .TextMatrix(lngCount, mColDetail.���ʱ��) = Format(.TextMatrix(lngCount, mColDetail.���ʱ��), "MM-dd hh:mm")
        Next
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize mColDetail.�¼���ʱ��
    End With
    If vfgThis.Rows > 1 Then vfgThis.Row = 1: Call vfgThis_RowColChange
End Sub

Private Sub vfgThis_RowColChange()
    On Error Resume Next
    Dim lngFileID As Long
    lngFileID = Val(vfgThis.TextMatrix(vfgThis.Row, mColDetail.��ɼ�¼id))
    If lngFileID > 0 Then
        mfrmContent.zlRefresh lngFileID
    Else
        mfrmContent.edtThis.ForceEdit = True
        mfrmContent.edtThis.ReadOnly = False
        mfrmContent.edtThis.NewDoc
        mfrmContent.edtThis.ReadOnly = True
        mfrmContent.edtThis.ForceEdit = False
    End If
End Sub


