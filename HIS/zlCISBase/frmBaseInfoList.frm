VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmBaseInfoList 
   Caption         =   "�������ݹ���"
   ClientHeight    =   7530
   ClientLeft      =   165
   ClientTop       =   870
   ClientWidth     =   11670
   Icon            =   "frmBaseInfoList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   11670
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picDesc 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   5160
      ScaleHeight     =   2295
      ScaleWidth      =   3135
      TabIndex        =   5
      Top             =   3720
      Width           =   3135
      Begin VB.Image imgNote 
         Height          =   240
         Left            =   120
         Picture         =   "frmBaseInfoList.frx":058A
         Top             =   240
         Width           =   240
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"frmBaseInfoList.frx":0B14
         ForeColor       =   &H00008000&
         Height          =   9000
         Left            =   360
         TabIndex        =   6
         Top             =   240
         Width           =   2460
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picType 
      BackColor       =   &H00FFEBD7&
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   5160
      ScaleHeight     =   3375
      ScaleWidth      =   3135
      TabIndex        =   3
      Top             =   360
      Width           =   3135
      Begin XtremeSuiteControls.ShortcutBar sbType 
         Height          =   3255
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   3015
         _Version        =   589884
         _ExtentX        =   5318
         _ExtentY        =   5741
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picList 
      BackColor       =   &H00FFEBD7&
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   720
      ScaleHeight     =   5295
      ScaleWidth      =   4425
      TabIndex        =   0
      Top             =   360
      Width           =   4425
      Begin XtremeReportControl.ReportControl rptList 
         Height          =   4410
         Left            =   480
         TabIndex        =   1
         Top             =   600
         Width           =   4395
         _Version        =   589884
         _ExtentX        =   7752
         _ExtentY        =   7779
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
      End
      Begin MSComctlLib.ImageList imgList 
         Left            =   0
         Top             =   4680
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
               Picture         =   "frmBaseInfoList.frx":0F10
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
      Top             =   7155
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmBaseInfoList.frx":14AA
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15505
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
   Begin MSComctlLib.ImageList ils32 
      Left            =   0
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseInfoList.frx":1D3C
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseInfoList.frx":2194
            Key             =   "Write"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseInfoList.frx":25E6
            Key             =   "Read"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   0
      Top             =   1140
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseInfoList.frx":2900
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseInfoList.frx":2D58
            Key             =   "Write"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseInfoList.frx":31AA
            Key             =   "Read"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   900
      Left            =   720
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5760
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
      Bindings        =   "frmBaseInfoList.frx":34C4
      Left            =   945
      Top             =   105
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmBaseInfoList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const conPane_Type = 201
Const conPane_List = 202
Const conPane_Edit = 203
Const conPane_Desc = 204
'-----------------------------------------------------
'�������
'-----------------------------------------------------
Private mstrPrivs As String     '��ǰʹ����Ȩ�޴�
Private mfrmEdit As frmBaseInfoEdit

Private mintEditState As Integer    '��ǰ�༭״̬��0-�Ǳ༭״̬,1-�༭״̬
Private mstr���� As String

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
Public Function zlRefList(strItemName As String) As Long
    '���ܣ�ˢ��װ��ָ������Ĳ����ļ��嵥������λ��ָ�����ļ���
    Dim i As Integer
    Dim rsTemp As New ADODB.Recordset
        
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select * from " & strItemName & " order by to_number(����)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Me.rptList.Records.DeleteAll
    With rsTemp
        Do While Not .EOF
            Set rptRcd = Me.rptList.Records.Add()
            For i = 0 To .Fields.Count - 1
                If i = 0 Then
                    Set rptItem = rptRcd.AddItem(CStr(Nvl(.Fields(i)))): rptItem.Icon = 0
                End If
                If .Fields(i).Name = "ȱʡ��־" Then
                    Set rptItem = rptRcd.AddItem(IIf(CStr(Nvl(.Fields(i))) = 1, "��", "")): rptItem.SortPriority = Val(("" & Nvl(.Fields(i))))
                Else
                    Set rptItem = rptRcd.AddItem(CStr(Nvl(.Fields(i))))   ': rptItem.SortPriority = Val(("" & Nvl(.Fields(i))))
                End If
            Next
            .MoveNext
        Loop
    End With
    With Me.rptList
        .GroupsOrder.DeleteAll
        .Populate
    End With

    If mstr���� <> "" Then
        For Each rptRow In Me.rptList.Rows
            If rptRow.GroupRow = False Then
                If Val(rptRow.Record(1).Value) = mstr���� Then
                    Set Me.rptList.FocusedRow = rptRow
                    Exit For
                End If
            End If
        Next
    End If
    If Me.rptList.FocusedRow Is Nothing And Me.rptList.Rows.Count > 0 Then
        If Me.rptList.Rows(0).GroupRow Then
            Set Me.rptList.FocusedRow = Me.rptList.Rows(0).Childs(0)
        Else
            Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
        End If
    End If
    Call rptList_SelectionChanged

    zlRefList = Me.rptList.Records.Count
    Me.stbThis.Panels(2).Text = "��" & strItemName & "������" & Me.rptList.Records.Count & "����¼"
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
    objPrint.Title.Text = "�����ʿع���"
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
    Dim str���� As String
    Dim lngRetuId As Long
    Dim panThis As Pane
    '------------------------------------
    Select Case Control.ID
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview: Call zlRptPrint(0)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_File_Exit: Unload Me: Unload frmMedTech: Unload frmMedTreat

    Case conMenu_Edit_Save:
        str���� = mfrmEdit.zlEditSave(gstrItemName)
        If str���� <> "" Then
            ShowEdit False
            mstr���� = str����: Call zlRefList(gstrItemName)
            mintEditState = 0: Me.picList.Enabled = True: Me.rptList.SetFocus
        End If
    Case conMenu_Edit_Untread:
        ShowEdit False
        Call mfrmEdit.zlEditCancel
        mintEditState = 0: Me.picList.Enabled = True: Me.rptList.SetFocus
        
    Case conMenu_Edit_NewItem
        mfrmEdit.fraEdit.BackColor = vbWhite
        ShowEdit True
        
        If mstr���� = "" Then Exit Sub
        If mfrmEdit.zlEditStart(True, gstrItemName, mstr����) = False Then Exit Sub
        mintEditState = 1: Me.picList.Enabled = False
        
    Case conMenu_Edit_Modify
        mfrmEdit.fraEdit.BackColor = vbWhite
        ShowEdit True
        
        If mstr���� = "" Then Exit Sub
        If mfrmEdit.zlEditStart(False, gstrItemName, mstr����) = False Then Exit Sub
        mintEditState = 1: Me.picList.Enabled = False
        
    Case conMenu_Edit_Delete
        Dim strMsg As String
        With Me.rptList
            strMsg = "���ɾ������Ŀ��¼��" & vbCrLf & "����" & .FocusedRow.Record(2).Value
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub

            gstrSql = "zl_" & gstrItemName & "_Edit(3,NULL,'" & mstr���� & "')"

            Err = 0: On Error GoTo ErrHand
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)

            Err = 0: On Error GoTo 0
            mstr���� = "": lngRetuId = .FocusedRow.Index
            If .Rows.Count > lngRetuId + 1 Then
                lngRetuId = lngRetuId + 1
            ElseIf lngRetuId > 0 Then
                lngRetuId = lngRetuId - 1
            End If
            If .Rows(lngRetuId).GroupRow = False Then mstr���� = .Rows(lngRetuId).Record(1).Value
            Call Me.zlRefList(gstrItemName)
        End With
'
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
        Call zlRefList(gstrItemName)

    Case conMenu_Help_Help:     Call ShowHelp(gstrLisHelp, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    End Select

    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
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
        Control.Enabled = (InStr(1, mstrPrivs, "��ɾ��") > 0 And mintEditState = 0 And Me.rptList.Rows.Count)
        If Control.Enabled Then Control.Enabled = mstr���� <> ""
        If Control.Enabled Then Control.Enabled = Not Me.rptList.FocusedRow.GroupRow
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    Case conMenu_View_Find, conMenu_View_Refresh, conMenu_View_Option: Control.Enabled = (mintEditState = 0)
    End Select
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_Type
        Item.Handle = picType.hWnd
    Case conPane_List
        Item.Handle = Me.picList.hWnd
    Case conPane_Edit
        If mfrmEdit Is Nothing Then Set mfrmEdit = New frmBaseInfoEdit
        Item.Handle = mfrmEdit.fraEdit.hWnd
    Case conPane_Desc
        Item.Handle = Me.picDesc.hWnd
    End Select
End Sub

Private Sub Form_Load()
    '-----------------------------------------------------
    'Ȩ�����ƴ����ƣ�����ͬʱ��������ģ�������gstrPrivs�仯�����¿�����Ч
    mstrPrivs = gstrPrivs
    
    mintEditState = 0
    mstr���� = 0
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, False)
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
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
    Dim panType As Pane, panList As Pane, panEdit As Pane, panDesc As Pane
    If mfrmEdit Is Nothing Then Set mfrmEdit = New frmBaseInfoEdit

    Set panType = dkpMan.CreatePane(conPane_Type, 160, 1000, DockLeftOf, Nothing)
    panType.Title = "������Ϣ����"
    panType.Options = PaneNoCaption Or PaneNoHideable Or PaneNoCloseable Or PaneNoFloatable
    
    Set panList = dkpMan.CreatePane(conPane_List, 800, 800, DockRightOf, panType)
    panList.Title = "������Ϣ�б�"
    panList.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set panEdit = dkpMan.CreatePane(conPane_Edit, 800, 200, DockBottomOf, panList)
    panEdit.Title = "������Ϣ�༭"
    panEdit.Options = PaneNoCaption
    panEdit.Close

    Set panDesc = dkpMan.CreatePane(conPane_Desc, 200, 1000, DockRightOf, Nothing)
    panDesc.Title = "����˵��"
    panDesc.Options = PaneNoCloseable

    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
    Call CreateShortCutBar
    
    Call DrawRpt(gstrItemName)          '��̬���� rptControl
    Call zlRefList(gstrItemName)        '����װ��
    Call LoadControl(gstrItemName)      '��̬���ر༭�ؼ�

    '����ָ�
    '-----------------------------------------------------
    Call RestoreWinState(Me, App.ProductName)
    Exit Sub
ErrHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload mfrmEdit
    Set mfrmEdit = Nothing
    Call SaveWinState(Me, App.ProductName)
    Unload Me
End Sub

Private Sub picList_Resize()
    Err = 0: On Error Resume Next
    With Me.rptList
        .Left = Me.picList.ScaleLeft: .Width = Me.picList.ScaleWidth - .Left
        .Top = Me.picList.ScaleTop: .Height = Me.picList.ScaleHeight - .Top
    End With
    mfrmEdit.fraEdit.Width = mfrmEdit.Width
    mfrmEdit.fraEdit.Height = Me.ScaleHeight - Me.picList.ScaleHeight
End Sub

Private Sub picType_Resize()
    Err = 0: On Error Resume Next
    With Me.sbType
        .Left = Me.picType.ScaleLeft: .Width = Me.picType.ScaleWidth - .Left
        .Top = Me.picType.ScaleTop: .Height = Me.picType.ScaleHeight - .Top
    End With
End Sub

Private Sub rptList_KeyDown(KeyCode As Integer, Shift As Integer)
    If Me.rptList.Visible = False Then Exit Sub
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Me.rptList.FocusedRow Is Nothing Then Exit Sub
    If Me.rptList.FocusedRow.GroupRow Then Exit Sub
    Call rptList_RowDblClick(Me.rptList.FocusedRow, Me.rptList.FocusedRow.Record.Item(0))
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
    If mstr���� = 0 Then Exit Sub

    Set cbrControl = Me.cbsThis.FindControl(, conMenu_Edit_Modify)
    If cbrControl Is Nothing Then Exit Sub
    If cbrControl.Visible = False Or cbrControl.Enabled = False Then Exit Sub
    Call cbsThis_Execute(cbrControl)

End Sub

Private Sub rptList_SelectionChanged()
    With Me.rptList
        If .FocusedRow Is Nothing Then
            mstr���� = 0
        ElseIf .FocusedRow.GroupRow = True Then
            mstr���� = 0
        Else
            mstr���� = Trim(.FocusedRow.Record.Item(0).Value)
        End If
        Call mfrmEdit.zlRefresh(gstrItemName, mstr����)
    End With
End Sub

Private Sub CreateShortCutBar()
    Dim objItem As ShortcutBarItem
    Dim objItemMain As ShortcutBarItem
      
    Set objItemMain = sbType.AddItem(1, "ҽ������", frmMedTreat.hWnd)
    Set objItem = sbType.AddItem(2, "ҽ�ƹ���", frmMedTech.hWnd)
    
    sbType.Selected = objItemMain
    sbType.ExpandedLinesCount = sbType.ItemCount
End Sub

Private Sub DrawRpt(strItemName As String)
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer
    
    On Err GoTo ErrHand:
    
    gstrSql = "select * from " & strItemName & " where rownum = 1 "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)

    rptList.Columns.DeleteAll
    rptList.AutoColumnSizing = (Screen.Width / Screen.TwipsPerPixelX > 800)   '������������֮ǰ���ã�������Ч
    
    Set rptCol = rptList.Columns.Add(0, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: _
                rptCol.Alignment = xtpAlignmentCenter
                
    For i = 0 To rsTemp.Fields.Count - 1
        Set rptCol = rptList.Columns.Add(i + 1, "" & rsTemp.Fields(i).Name, 85, True): rptCol.Editable = False: rptCol.Groupable = False
    Next
    

    rptList.SetImageList Me.ils16
    rptList.AllowColumnRemove = False
    rptList.MultipleSelection = False
    rptList.ShowItemsInGroups = False
    With rptList.PaintManager
        .ColumnStyle = xtpColumnFlat
        .GridLineColor = RGB(225, 225, 225)
        .NoGroupByText = "�϶��б��⵽����,����������..."
        .NoItemsText = "û�п���ʾ����Ŀ..."
        .VerticalGridStyle = xtpGridSolid
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub ShowItemInfo(strItemName As String)
    gstrItemName = strItemName
    Call DrawRpt(gstrItemName)
    Call zlRefList(gstrItemName)
    Call LoadControl(gstrItemName)
End Sub

Private Sub LoadControl(strItemName As String)
    Dim objControl As Object
    Dim int��� As Integer: int��� = 570
    Dim intLblAndTxt As Integer: intLblAndTxt = 210

    '���������صĿؼ�ȫ����ʾ
    For Each objControl In mfrmEdit.Controls
        If objControl.Visible = False Then
            objControl.Visible = True
        End If
    Next

    '�ָ��ؼ�ԭ�д�С
    mfrmEdit.txt����.Left = 1050
    mfrmEdit.txt����.Top = 360
    mfrmEdit.txt����.Width = 1380

    mfrmEdit.txt����.Left = 3570
    mfrmEdit.txt����.Top = 360
    mfrmEdit.txt����.Width = 2235

    mfrmEdit.cbo�����Ա�.Left = 7035
    mfrmEdit.cbo�����Ա�.Top = 360
    mfrmEdit.cbo�����Ա�.Width = 2235

    mfrmEdit.txt����.Left = 7035
    mfrmEdit.txt����.Top = 360
    mfrmEdit.txt����.Width = 1215

    mfrmEdit.txt˵��.Left = 1050
    mfrmEdit.txt˵��.Top = 840
    mfrmEdit.txt˵��.Width = 3285
    mfrmEdit.txt˵��.Height = 720

    mfrmEdit.txt����.Left = 1050
    mfrmEdit.txt����.Top = 840
    mfrmEdit.txt����.Width = 615

    mfrmEdit.chkȱʡ��־.Left = 4980
    mfrmEdit.chkȱʡ��־.Top = 840
    mfrmEdit.chkȱʡ��־.Width = 3255

    mfrmEdit.cbo����.Left = 10290
    mfrmEdit.cbo����.Top = 360
    mfrmEdit.cbo����.Width = 1395
    
    mfrmEdit.cbo����1.Left = 10290
    mfrmEdit.cbo����1.Top = 360
    mfrmEdit.cbo����1.Width = 1395
    mfrmEdit.cbo����1.Visible = False
    
    mfrmEdit.lbl1.Left = 6375
    mfrmEdit.lbl1.Top = 420
    mfrmEdit.lbl˵��.Caption = "˵��"

    '���ݲ�ͬ��ѡ�����¼��ؽ���ؼ�
    With mfrmEdit
        Select Case Trim(strItemName)
        Case "���Ƽ���걾"
            .txt˵��.Visible = False
            .txt����.Visible = False
            .chkȱʡ��־.Visible = False
            .lbl˵��.Visible = False
            .cbo����.Visible = False
            .lbl1.Caption = "����"
            .txt����.Left = .cbo�����Ա�.Left
            .txt����.Top = .cbo�����Ա�.Top
            .lbl����.Left = .txt����.Left + .txt����.Width + int���
            .lbl����.Caption = "�����Ա�"
            .lbl����.Width = 2 * .lbl����.Width
            .cbo�����Ա�.Left = .lbl����.Left + .lbl����.Width + intLblAndTxt
            .cbo�����Ա�.Top = .cbo����.Top
            .cbo�����Ա�.Width = 800
            
            .txt����.MaxLength = 2
            .txt����.MaxLength = 20
            .txt����.MaxLength = 8
            
        Case "���Ƽ�������"
            .txt˵��.Visible = False
            .lbl����.Visible = False
            .cbo����.Visible = False
            .cbo�����Ա�.Visible = False
            .lbl1.Caption = "����"
            .txt����.Left = .cbo�����Ա�.Left
            .txt����.Top = .cbo�����Ա�.Top
            .lbl˵��.Caption = "����"
            .txt����.Left = .txt˵��.Left
            .txt����.Top = .txt˵��.Top
            .chkȱʡ��־.Left = .lbl����.Left
            
            .txt����.MaxLength = 2
            .txt����.MaxLength = 20
            .txt����.MaxLength = 8
            .txt����.MaxLength = 2
            
        Case "���鱸ע����"
            .cbo�����Ա�.Visible = False
            .txt����.Visible = False
            .chkȱʡ��־.Visible = False
            .lbl1.Caption = "����"
            .lbl˵��.Caption = "˵��"
            .txt����.Left = .cbo�����Ա�.Left
            .txt����.Top = .cbo�����Ա�.Top
            .lbl����.Caption = "����"
            .cbo����.Left = .lbl����.Left + .lbl����.Width + intLblAndTxt
            
            .txt����.MaxLength = 10
            .txt����.MaxLength = 100
            .txt����.MaxLength = 10
            .txt˵��.MaxLength = 80

        Case "������������"
            .txt����.Visible = False
            .chkȱʡ��־.Visible = False
            .lbl����.Visible = False
            .cbo����.Visible = False
            .cbo�����Ա�.Visible = False
            .txt����.Width = .txt����.Width * 2
            .lbl1.Caption = "����"
            .lbl1.Left = .txt����.Left + .txt����.Width + int���
            .txt����.Left = .lbl1.Left + .lbl1.Width + intLblAndTxt
            .txt����.Top = .cbo�����Ա�.Top
            .lbl˵��.Caption = "˵��"
            
            .txt����.MaxLength = 10
            .txt����.MaxLength = 100
            .txt����.MaxLength = 10
            .txt˵��.MaxLength = 80

        Case "������������"
            .cbo�����Ա�.Visible = False
            .txt����.Visible = False
            .chkȱʡ��־.Visible = False
            .lbl1.Caption = "����"
            .lbl����.Caption = "����"
            .txt����.Left = .cbo�����Ա�.Left
            .txt����.Top = .cbo�����Ա�.Top
            .lbl˵��.Caption = "˵��"
            .cbo����.Left = .lbl����.Left + .lbl����.Width + intLblAndTxt
            
            .txt����.MaxLength = 3
            .txt����.MaxLength = 50
            .txt����.MaxLength = 10
            .txt˵��.MaxLength = 80
            
        Case "����걾��̬"
            .cbo�����Ա�.Visible = False
            .lbl1.Visible = False
            .txt����.Visible = False
            .txt����.Visible = False
            .cbo����.Visible = False
            .chkȱʡ��־.Visible = False
            .lbl����.Visible = False
            
            .txt����.MaxLength = 10
            .txt����.MaxLength = 50
            .txt˵��.MaxLength = 100

        Case "���������;"
            .txt˵��.Visible = False
            .lbl1.Visible = False
            .cbo�����Ա�.Visible = False
            .txt����.Visible = False
            .txt����.Visible = False
            .cbo����.Visible = False
            .chkȱʡ��־.Visible = False
            .lbl˵��.Visible = False
            .lbl����.Visible = False
            .txt����.Width = .txt����.Width * 2
            
            .txt����.MaxLength = 10
            .txt����.MaxLength = 200
            
        Case "�����������"
            .lbl1.Visible = False
            .cbo�����Ա�.Visible = False
            .txt����.Visible = False
            .txt����.Visible = False
            .cbo����.Visible = False
            .chkȱʡ��־.Visible = False
            .lbl����.Visible = False
            .lbl����.Visible = False
            .txt����.Visible = False
            .lbl˵��.Caption = "����"
            
            .txt����.MaxLength = 10
            .txt˵��.MaxLength = 200
            
        Case "����������"
            .cbo�����Ա�.Visible = False
            .lbl˵��.Visible = False
            .txt˵��.Visible = False
            .txt����.Visible = False
            .lbl����.Visible = False
            .cbo����.Visible = False
            .lbl1.Caption = "����"
            .txt����.Left = .cbo�����Ա�.Left
            .txt����.Top = .cbo�����Ա�.Top
            .chkȱʡ��־.Left = .lbl˵��.Left
            .chkȱʡ��־.Top = .lbl˵��.Top
            
            .txt����.MaxLength = 2
            .txt����.MaxLength = 20
            .txt����.MaxLength = 8

        Case "����ϸ�����"
            .cbo�����Ա�.Visible = False
            .lbl˵��.Visible = False
            .txt˵��.Visible = False
            .txt����.Visible = False
            .lbl����.Visible = False
            .cbo����.Visible = False
            .lbl1.Caption = "����"
            .txt����.Left = .cbo�����Ա�.Left
            .txt����.Top = .cbo�����Ա�.Top
            .chkȱʡ��־.Left = .lbl˵��.Left
            .chkȱʡ��־.Top = .lbl˵��.Top
            
            .txt����.MaxLength = 8
            .txt����.MaxLength = 30
            .txt����.MaxLength = 20

        Case "����ϸ������"
            .cbo�����Ա�.Visible = False
            .lbl˵��.Visible = False
            .txt˵��.Visible = False
            .txt����.Visible = False
            .chkȱʡ��־.Visible = False
            .lbl����.Visible = False
            .cbo����.Visible = False
            .lbl1.Caption = "����"
            .txt����.Left = .cbo�����Ա�.Left
            .txt����.Top = .cbo�����Ա�.Top
            
            .txt����.MaxLength = 8
            .txt����.MaxLength = 30
            .txt����.MaxLength = 20

        Case "����Ⱦɫ����"
            .cbo�����Ա�.Visible = False
            .lbl˵��.Visible = False
            .txt˵��.Visible = False
            .txt����.Visible = False
            .lbl����.Visible = False
            .cbo����.Visible = False
            .lbl1.Caption = "����"
            .txt����.Left = .cbo�����Ա�.Left
            .txt����.Top = .cbo�����Ա�.Top
            .chkȱʡ��־.Left = .lbl˵��.Left
            .chkȱʡ��־.Top = .lbl˵��.Top
            
            .txt����.MaxLength = 8
            .txt����.MaxLength = 30
            .txt����.MaxLength = 20

        Case "�ʿر���ʾ�"
            .txt˵��.Visible = False
            .txt����.Visible = False
            .chkȱʡ��־.Visible = False
            .cbo����.Visible = False
            .txt����.Width = .txt����.Width * 3
            .lbl˵��.Caption = "����"
            .txt����.Left = .txt˵��.Left
            .txt����.Top = .txt˵��.Top
            .lbl1.Caption = "����"
            .lbl1.Left = .lbl����.Left
            .lbl1.Top = .lbl˵��.Top
            .cbo�����Ա�.Left = .lbl1.Left + .lbl1.Width + intLblAndTxt
            .cbo�����Ա�.Top = .lbl1.Top
            .cbo�����Ա�.Width = 1000
            .cbo����1.Left = .cbo�����Ա�.Left
            .cbo����1.Top = .cbo�����Ա�.Top
            .cbo�����Ա�.Width = .cbo�����Ա�.Width
            .cbo����1.Visible = True
            .cbo�����Ա�.Visible = False
            
            .txt����.MaxLength = 3
            .txt����.MaxLength = 80
            .txt����.MaxLength = 10
            
        Case "�ʿؼ��鷽��"
            .cbo�����Ա�.Visible = False
            .lbl˵��.Visible = False
            .txt˵��.Visible = False
            .txt����.Visible = False
            .chkȱʡ��־.Visible = False
            .lbl����.Visible = False
            .cbo����.Visible = False
            .lbl1.Caption = "����"
            .txt����.Left = .cbo�����Ա�.Left
            .txt����.Top = .cbo�����Ա�.Top
            
            .txt����.MaxLength = 6
            .txt����.MaxLength = 30
            .txt����.MaxLength = 10

        Case "�ʿ��Լ���Դ"
            .cbo�����Ա�.Visible = False
            .txt˵��.Visible = False
            .chkȱʡ��־.Visible = False
            .lbl˵��.Visible = False
            .cbo����.Visible = False
            .lbl1.Caption = "����"
            .txt����.Left = .cbo�����Ա�.Left
            .txt����.Top = .cbo�����Ա�.Top
            .lbl����.Caption = "QC����"
            .txt����.Left = .lbl����.Left + .lbl����.Width + intLblAndTxt
            .txt����.Top = .cbo����.Top
            .txt����.Width = 800
            
            .txt����.MaxLength = 6
            .txt����.MaxLength = 30
            .txt����.MaxLength = 10
            .txt����.MaxLength = 8

        Case "����������"
            .txt˵��.Visible = False
            .txt����.Visible = False
            .chkȱʡ��־.Visible = False
            .lbl˵��.Visible = False
            .cbo�����Ա�.Visible = False
            .lbl1.Caption = "����"
            .txt����.Left = .cbo�����Ա�.Left
            .txt����.Top = .cbo�����Ա�.Top
            .cbo����.Width = 1000
            .cbo����1.Width = .cbo����.Width
            .cbo����.Left = .lbl����.Left + .lbl����.Width + intLblAndTxt
            .cbo����1.Left = .cbo����.Left
            .cbo����.Visible = False
            .cbo����1.Visible = True
            .lbl����.Caption = "����"
            
            .txt����.MaxLength = 3
            .txt����.MaxLength = 200
            .txt����.MaxLength = 20
            
            
        Case "ϸ����ⷽ��"
            .cbo�����Ա�.Visible = False
            .lbl˵��.Visible = False
            .txt˵��.Visible = False
            .txt����.Visible = False
            .chkȱʡ��־.Visible = False
            .lbl����.Visible = False
            .cbo����.Visible = False
            .lbl1.Caption = "����"
            .txt����.Left = .cbo�����Ա�.Left
            .txt����.Top = .cbo�����Ա�.Top
            
            .txt����.MaxLength = 2
            .txt����.MaxLength = 20
            .txt����.MaxLength = 10
            
        Case "ϸ����ҩ����"
            .cbo�����Ա�.Visible = False
            .lbl˵��.Visible = False
            .txt˵��.Visible = False
            .txt����.Visible = False
            .chkȱʡ��־.Visible = False
            .lbl����.Visible = False
            .cbo����.Visible = False
            .lbl1.Caption = "����"
            .txt����.Left = .cbo�����Ա�.Left
            .txt����.Top = .cbo�����Ա�.Top
            
            .txt����.MaxLength = 4
            .txt����.MaxLength = 100
            .txt����.MaxLength = 20
            
        End Select
    End With
End Sub

Private Sub sbType_ExpandButtonDown(CancelMenu As Boolean)
    CancelMenu = True
End Sub

Private Sub sbType_SelectedChanged(ByVal Item As XtremeSuiteControls.IShortcutBarItem)
    Dim i As Integer
    Select Case Item.ID
        Case "1"
            For i = 1 To frmMedTreat.tplFunc.Groups(1).Items.Count
                If frmMedTreat.tplFunc.Groups(1).Items(i).Selected Then
                    Call ShowItemInfo(frmMedTreat.tplFunc.Groups(1).Items(i).Caption)
                    Exit For
                End If
            Next
            
        Case "2"
            For i = 1 To frmMedTech.tplFunc.Groups(1).Items.Count - 1
                If frmMedTech.tplFunc.Groups(1).Items(i).Selected Then
                    Call ShowItemInfo(frmMedTech.tplFunc.Groups(1).Items(i).Caption)
                    Exit For
                End If
            Next
    End Select
End Sub

Private Sub ShowEdit(blnShow As Boolean)
    '����       �Ƿ���ʾ�ǼǴ���
    Dim Pane1 As Pane
    Set Pane1 = dkpMan.FindPane(conPane_Edit)
    If blnShow = True Then
        Pane1.Select
    Else
        Pane1.Close
    End If
    dkpMan.RecalcLayout
End Sub




