VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Begin VB.Form frmLabRptList 
   Caption         =   "����ģ�����"
   ClientHeight    =   6435
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   8970
   Icon            =   "frmLabRptList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picList 
      BackColor       =   &H00FFEBD7&
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   75
      ScaleHeight     =   5295
      ScaleWidth      =   4425
      TabIndex        =   0
      Top             =   450
      Width           =   4425
      Begin XtremeReportControl.ReportControl rptList 
         Height          =   4410
         Left            =   30
         TabIndex        =   1
         Top             =   60
         Width           =   4350
         _Version        =   589884
         _ExtentX        =   7673
         _ExtentY        =   7779
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
      End
      Begin MSComctlLib.ImageList imgList 
         Left            =   90
         Top             =   4800
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLabRptList.frx":058A
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   6060
      Width           =   8970
      _ExtentX        =   15822
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmLabRptList.frx":0B24
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10742
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
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   900
      Left            =   180
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5625
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmLabRptList.frx":13B6
      Left            =   945
      Top             =   105
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmLabRptList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mCol
    ͼ�� = 0: ID: ����: ����: ����: ������Ŀid: ������Ŀ��: ˵��
End Enum

Const conPane_List = 201
Const conPane_Edit = 202

'-----------------------------------------------------
'�������
'-----------------------------------------------------
Private mstrPrivs As String     '��ǰʹ����Ȩ�޴�

Private mLngEditWidth As Long       'Ϊ��Ӧ����������´�����.�ȶ��봰���С.

Private mfrmEdit As frmLabRptEdit

Private mintEditState As Integer    '��ǰ�༭״̬��0-�Ǳ༭״̬,1-�༭״̬
Private mlngItemID As Long

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

Dim lngCount As Long

'-----------------------------------------------------
'����Ϊ�ڲ���������
'-----------------------------------------------------
Public Function zlRefList(Optional lngItemID As Long) As Long
    '���ܣ�ˢ��װ��ָ������Ĳ����ļ��嵥������λ��ָ�����ļ���
    Dim rsTemp As New ADODB.Recordset
    gstrSql = "Select L.ID, L.����, L.����, L.����, L.������Ŀid, I.���� As ������Ŀ��, L.˵��" & vbNewLine & _
            "From ����ģ��Ŀ¼ L, ������ĿĿ¼ I" & vbNewLine & _
            "Where L.������Ŀid = I.ID"
    
    Err = 0: On Error GoTo ErrHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Me.rptList.Records.DeleteAll
    With rsTemp
        Do While Not .EOF
            Set rptRcd = Me.rptList.Records.Add()
            Set rptItem = rptRcd.AddItem(Val(0)): rptItem.Icon = rptItem.Value
            rptRcd.AddItem CStr("" & !ID)
            Set rptItem = rptRcd.AddItem(CStr("" & !����)): rptItem.SortPriority = Val(("" & !����))
            rptRcd.AddItem CStr("" & !����)
            rptRcd.AddItem CStr("" & !����)
            rptRcd.AddItem CStr("" & !������Ŀid)
            rptRcd.AddItem CStr("" & !������Ŀ��)
            rptRcd.AddItem CStr("" & !˵��)
            .MoveNext
        Loop
    End With
    Me.rptList.Populate
    
    If lngItemID <> 0 Then
        For Each rptRow In Me.rptList.Rows
            If rptRow.GroupRow = False Then
                If Val(rptRow.Record(mCol.ID).Value) = lngItemID Then
                    Set Me.rptList.FocusedRow = rptRow
                    mlngItemID = Val(Me.rptList.FocusedRow.Record(mCol.ID).Value)
                    Exit For
                End If
            End If
        Next
    End If
    If Me.rptList.FocusedRow Is Nothing Then
        If Me.rptList.Rows.Count > 0 Then
            Set Me.rptList.FocusedRow = Me.rptList.Rows(1)
            mlngItemID = Val(Me.rptList.FocusedRow.Record(mCol.ID).Value)
        Else
            mlngItemID = 0: Call mfrmEdit.zlRefresh(mlngItemID)
        End If
    End If
    
    zlRefList = Me.rptList.Records.Count
    Me.stbThis.Panels(2).Text = "����" & Me.rptList.Records.Count & "����¼"
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefList = Me.rptList.Records.Count
End Function

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode��1-��ӡ;2-Ԥ��;3-�����EXCEL
    If Me.rptList.Records.Count = 0 Then Exit Sub
    
    '-------------------------------------------------
    '�������ݱ��
    If zlControl.RPTCopyToVSF(Me.rptList, Me.vfgList) Is Nothing Then Exit Sub
    '-------------------------------------------------
    '���ô�ӡ��������
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    Set objPrint.Body = Me.vfgList
    objPrint.Title.Text = "���鱨��ģ���嵥"
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

'-----------------------------------------------------
'����Ϊ�ؼ��¼�����
'-----------------------------------------------------
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRetuId As Long
    
    '------------------------------------
    Select Case Control.ID
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview: Call zlRptPrint(0)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_File_Exit: Unload Me
    
    Case conMenu_Edit_Save:
        lngRetuId = mfrmEdit.zlEditSave()
        If lngRetuId <> 0 Then
            mlngItemID = lngRetuId
            If mintEditState = 1 Then
                Call zlRefList(mlngItemID)
            Else
                Call zlRefList(mlngItemID)
            End If
            mintEditState = 0: Me.picList.Enabled = True
        End If
    Case conMenu_Edit_Untread:
        Call mfrmEdit.zlEditCancel
        Call zlRefList(mlngItemID)
        mintEditState = 0: Me.picList.Enabled = True
    Case conMenu_Edit_NewItem
        If mfrmEdit.zlEditStart(True, mlngItemID) Then mintEditState = 1: Me.picList.Enabled = False
        Me.dkpMan.FindPane(conPane_Edit).Select

    Case conMenu_Edit_Modify
        If mlngItemID = 0 Then Exit Sub
        If mfrmEdit.zlEditStart(False, mlngItemID) Then mintEditState = 1
        Me.picList.Enabled = False
        Me.dkpMan.FindPane(conPane_Edit).Select

    Case conMenu_Edit_Delete
        Dim strMsg As String
        With Me.rptList
            If .FocusedRow.GroupRow = True Then Exit Sub
            strMsg = "���ɾ������ģ����" & vbCrLf & "����" & .FocusedRow.Record(mCol.����).Value
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            gstrSql = "Zl_����ģ��Ŀ¼_Edit(3," & mlngItemID & ")"
            
            Err = 0: On Error GoTo ErrHand
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)

            Err = 0: On Error GoTo 0
            mlngItemID = 0: lngRetuId = .FocusedRow.Index
            If .Rows.Count > lngRetuId + 1 Then
                If .Rows(lngRetuId + 1).GroupRow = False Then mlngItemID = .Rows(lngRetuId + 1).Record(mCol.ID).Value
            ElseIf lngRetuId > 0 Then
                If .Rows(lngRetuId - 1).GroupRow = False Then mlngItemID = .Rows(lngRetuId - 1).Record(mCol.ID).Value
            End If
            Call Me.zlRefList(mlngItemID)
        End With
    
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
        Call zlRefList(mlngItemID)
    
    Case conMenu_Help_Help:     Call ShowHelp(gstrLisHelp, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    End Select
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = (Me.rptList.Records.Count <> 0 And mintEditState = 0)
    Case conMenu_Edit_Save, conMenu_Edit_Untread
        Control.Enabled = (mintEditState <> 0)
    Case conMenu_Edit_NewItem
        Control.Enabled = (InStr(1, mstrPrivs, "��ɾ��") > 0 And mintEditState = 0)
    Case conMenu_Edit_Modify, conMenu_Edit_Delete
        Control.Enabled = (InStr(1, mstrPrivs, "��ɾ��") > 0 And mintEditState = 0 And mlngItemID <> 0)
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    Case conMenu_View_Find, conMenu_View_Refresh, conMenu_View_Option: Control.Enabled = (mintEditState = 0)
    End Select
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_List
        Item.Handle = Me.picList.hWnd
    Case conPane_Edit
        If mfrmEdit Is Nothing Then Set mfrmEdit = New frmLabRptEdit
        Item.Handle = mfrmEdit.hWnd
    End Select
End Sub

Private Sub Form_Load()
    '-----------------------------------------------------
    'Ȩ�����ƴ����ƣ�����ͬʱ��������ģ�������gstrPrivs�仯�����¿�����Ч
    mstrPrivs = gstrPrivs
    
    mLngEditWidth = frmLabRptEdit.ScaleWidth
    
    mintEditState = 0
    mlngItemID = 0
    
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, False)
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    'Me.cbsThis.Icons = frmPubIcons.imgPublic.Icons
    Set cbsThis.Icons = zlCommFun.GetPubIcons
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "����(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "ȡ��(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&D)")
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
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save
        .Add FCONTROL, Asc("Z"), conMenu_Edit_Untread
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FSHIFT, VK_DELETE, conMenu_Edit_Delete
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '���ò����ò˵�
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    '-----------------------------------------------------
    '����������
    Set cbrToolBar = Me.cbsThis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "ȡ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '-----------------------------------------------------
    '���ôʾ���ʾͣ������
    Dim panThis As Pane
    If mfrmEdit Is Nothing Then Set mfrmEdit = New frmLabRptEdit
    
    Set panThis = dkpMan.CreatePane(conPane_List, 450, 580, DockLeftOf, Nothing)
    panThis.Title = "ģ���б�"
    panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Set panThis = dkpMan.CreatePane(conPane_Edit, 550, 580, DockRightOf, Nothing)
    panThis.Title = "ģ����Ϣ"
    panThis.Options = PaneNoCaption

    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
    
    '-----------------------------------------------------
    With Me.rptList
        .AutoColumnSizing = (Screen.Width / Screen.TwipsPerPixelX > 800)   '������������֮ǰ���ã�������Ч
        Set rptCol = .Columns.Add(mCol.ͼ��, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mCol.ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.����, "����", 60, False): rptCol.Editable = False: rptCol.Groupable = False: .SortOrder.Add rptCol
        Set rptCol = .Columns.Add(mCol.����, "����", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.����, "����", 70, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.������Ŀid, "��Ŀid", 0, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.������Ŀ��, "��Ŀ��", 70, False): rptCol.Editable = False: rptCol.Groupable = True: rptCol.Visible = False
        .GroupsOrder.Add rptCol
        Set rptCol = .Columns.Add(mCol.˵��, "˵��", 240, True): rptCol.Editable = False: rptCol.Groupable = False
        
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
    Call zlRefList

End Sub

Private Sub Form_Resize()
    Dim panKind As Pane
    If Me.WindowState = vbMinimized Then Exit Sub
    Set panKind = Me.dkpMan.FindPane(conPane_Edit)
    panKind.MinTrackSize.SetSize mLngEditWidth / Screen.TwipsPerPixelX, panKind.MinTrackSize.Height
    panKind.MaxTrackSize.SetSize mLngEditWidth / Screen.TwipsPerPixelX, panKind.MaxTrackSize.Height
    Me.dkpMan.RecalcLayout
    Me.dkpMan.NormalizeSplitters
    panKind.MinTrackSize.SetSize 0, panKind.MinTrackSize.Height
    panKind.MaxTrackSize.SetSize mLngEditWidth / Screen.TwipsPerPixelX, panKind.MaxTrackSize.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload mfrmEdit
    Set mfrmEdit = Nothing
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub picList_Resize()
    With Me.rptList
        .Left = Me.picList.ScaleLeft: .Width = Me.picList.ScaleWidth - .Left
        .Top = Me.picList.ScaleTop: .Height = Me.picList.ScaleHeight - .Top
    End With
End Sub

Private Sub rptList_KeyDown(KeyCode As Integer, Shift As Integer)
    If Me.rptList.Visible = False Then Exit Sub
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Me.rptList.FocusedRow Is Nothing Then Exit Sub
    If Me.rptList.FocusedRow.GroupRow Then Exit Sub
    Call rptList_RowDblClick(Me.rptList.FocusedRow, Me.rptList.FocusedRow.Record.Item(mCol.ID))
End Sub

Private Sub rptList_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    
    If Button <> vbRightButton Then Exit Sub
    If Me.cbsThis.ActiveMenuBar.Controls(2).Visible = False Then Exit Sub

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls(2)
    Set cbrPopupBar = Me.cbsThis.Add("�����˵�", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, cbrControl.ID, cbrControl.Caption)
        cbrPopupItem.BeginGroup = cbrControl.BeginGroup
    Next
    cbrPopupBar.ShowPopup
End Sub

Private Sub rptList_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If mlngItemID = 0 Then Exit Sub
    
    Set cbrControl = Me.cbsThis.FindControl(, conMenu_Edit_Modify)
    If cbrControl Is Nothing Then Exit Sub
    If cbrControl.Visible = False Or cbrControl.Enabled = False Then Exit Sub
    Call cbsThis_Execute(cbrControl)

End Sub

Private Sub rptList_SelectionChanged()
    With Me.rptList
        If .FocusedRow Is Nothing Then
            mlngItemID = 0
        ElseIf .FocusedRow.GroupRow = True Then
            mlngItemID = 0
        Else
            mlngItemID = Val(.FocusedRow.Record.Item(mCol.ID).Value)
        End If
        Call mfrmEdit.zlRefresh(mlngItemID)
    End With
End Sub


