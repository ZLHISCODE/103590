VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEPRFileMan 
   Caption         =   "�����ļ�����"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8640
   Icon            =   "frmEPRFileMan.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   8640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox PicFileTab 
      BorderStyle     =   0  'None
      Height          =   5010
      Left            =   195
      ScaleHeight     =   5010
      ScaleWidth      =   4410
      TabIndex        =   2
      Top             =   660
      Width           =   4410
      Begin XtremeReportControl.ReportControl rptList 
         Height          =   4440
         Left            =   15
         TabIndex        =   3
         Top             =   510
         Width           =   3975
         _Version        =   589884
         _ExtentX        =   7011
         _ExtentY        =   7832
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
      End
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   870
         TabIndex        =   4
         Top             =   75
         Width           =   3105
      End
      Begin VB.Label lblFind 
         Caption         =   "����(&V)"
         Height          =   405
         Left            =   135
         TabIndex        =   5
         Top             =   105
         Width           =   945
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6015
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmEPRFileMan.frx":058A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12330
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
      Left            =   2730
      Top             =   60
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
            Picture         =   "frmEPRFileMan.frx":0E1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRFileMan.frx":13B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRFileMan.frx":1950
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRFileMan.frx":1EEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRFileMan.frx":2484
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRFileMan.frx":2A1E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vgdList 
      Height          =   900
      Left            =   300
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5730
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
      Left            =   2070
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   270
      Top             =   105
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmEPRFileMan.frx":2FB8
      Left            =   960
      Top             =   210
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmEPRFileMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------
'����
'-----------------------------------------------------
Private Enum mCol
    ͼ�� = 0: ID: ����: ���: ����: ˵��: ����: ҳ��: ����: ����
End Enum
Const conPane_FileTab = 1
Const conPane_Request = 2
Const conPane_Compend = 3

'-----------------------------------------------------
'�������
'-----------------------------------------------------
Private mstrPrivs As String     '��ǰʹ����Ȩ�޴�
Private mstrKinds As String     '��ǰ������Ĳ������ʹ�

Private WithEvents mfrmRequest As frmEPRFileRequest     'Ӧ��Ҫ�󴰸�
Attribute mfrmRequest.VB_VarHelpID = -1
Private WithEvents mfrmContent As frmEPRFileContent     '������ٴ���
Attribute mfrmContent.VB_VarHelpID = -1
Private mObjTabEpr As cTableEPR

Private mintCurKind As Integer      '��������
Private mlngCurFileId As Long       '��ǰ�ļ�ID
Private mstrCurFixed As String      '������������
Private mblnPartogram As Boolean    '�Ƿ��ǲ����ļ� (����=3 And ����=1)

Private mblnFindTag As Boolean      '�����򽹵��ж�
Private mintLastRows As Integer     '�������λ��λ��

Public Sub RefreshList()
    Call mfrmContent.zlRefresh(mlngCurFileId)
End Sub

Public Function zlRefList(Optional lngFileID As Long) As Long
    '���ܣ�ˢ��װ��ָ������Ĳ����ļ��嵥������λ��ָ�����ļ���
Dim strGroups As String
Dim rsTemp As New ADODB.Recordset
Dim rptRcd As ReportRecord
Dim rptItem As ReportRecordItem
Dim rptRow As ReportRow


    Me.rptList.Tag = ""
    gstrSQL = "Select l.Id, l.����, l.���, l.����, l.˵��, Nvl(l.����, 0) As ����, Decode(f.����, 1, '����ҳ��', f.����) As ҳ��,l.����" & _
            " From �����ļ��б� l," & _
            "      (Select f.����, f.���, f.����, Count(l.ID) As ����" & _
            "        From ����ҳ���ʽ f, �����ļ��б� l" & _
            "        Where f.���� = l.���� And f.��� = l.ҳ�� And f.���� In (" & mstrKinds & ")" & _
            "        Group By f.����, f.���, f.����) f" & _
            " Where l.���� = f.���� And l.ҳ�� = f.��� and l.����<>4"
    
    Err = 0: On Error GoTo errHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
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
            Case 5 And !���� <> 4: rptRcd.AddItem CStr("5-����֤������")
            Case 6: rptRcd.AddItem CStr("6-֪���ļ�")
            Case Else: rptRcd.AddItem ""
            End Select
            rptRcd.AddItem CStr(!���)
            rptRcd.AddItem CStr(!����)
            rptRcd.AddItem CStr("" & !˵��)
            Select Case !����
            Case 0: rptRcd.AddItem ""
            Case 1: rptRcd.AddItem CStr("����")
            Case 2: rptRcd.AddItem CStr("���")
            Case 3: rptRcd.AddItem CStr("���")
            Case Else
                If NVL(!����) = 3 And NVL(!����) = -1 Then
                    rptRcd.AddItem "����"
                Else
                    rptRcd.AddItem CStr("����")
                End If
            End Select
            rptRcd.AddItem CStr(!ҳ��)
            rptRcd.AddItem CStr(NVL(!����))
            rptRcd.AddItem zl9ComLib.zlStr.PinYinCode(CStr(!����))
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
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    zlRefList = Me.rptList.Records.Count
    lngFileID = 0
End Function

Public Sub zlRptPrint(ByVal bytMode As Byte)
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode��1-��ӡ;2-Ԥ��;3-�����EXCEL
    If Me.rptList.Records.Count = 0 Then Exit Sub

    '-------------------------------------------------
    '�������ݱ��
    If zlReportToVSFlexGrid(Me.vgdList, Me.rptList) = False Then Exit Sub

    '-------------------------------------------------
    '���ô�ӡ��������
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow

    Set objPrint.Body = Me.vgdList
    objPrint.Title.Text = "�����ļ��嵥"
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
    Dim lngFileID As Long, lngCopyId As Long
    Dim cbrControl As CommandBarControl
    Dim str��� As String, str���� As String
    
    Select Case Control.ID
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview: Call zlRptPrint(0)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_File_ExportToXML + 1
       frmFileExportOrImport.ShowMe Me, 1
    Case conMenu_File_ExportToXML + 2
        frmFileExportOrImport.ShowMe Me, 2
    Case conMenu_File_ExportToXML
        '������XML�ļ�
        If Me.rptList.FocusedRow Is Nothing Then Exit Sub
        If Me.rptList.FocusedRow.GroupRow = True Then Exit Sub
        Dim strF As String
        lngFileID = Me.rptList.FocusedRow.Record.Item(mCol.ID).Value
        'ָ��������ļ�·��
        On Error Resume Next
        dlgThis.Filename = "����_" & Me.rptList.FocusedRow.Record.Item(mCol.����).Value & ".xml"
        dlgThis.Filter = "*.XML|*.xml|*.*|*.*"
        dlgThis.CancelError = True
        dlgThis.ShowSave
        If Err.Number = 32755 Then Err.Clear: Exit Sub
        strF = dlgThis.Filename
        On Error GoTo errHand
        If gobjFSO.FileExists(strF) Then
            DoEvents
            If MsgBox("���ļ��Ѿ����ڣ��Ƿ񸲸ǣ�", vbOKCancel + vbQuestion, gstrSysName) = vbCancel Then Exit Sub
        End If
        
        If mstrCurFixed = "���" Then '���ʽ��������
            mObjTabEpr.InitOpenEPR Me, cprEM_�޸�, cprET_�����ļ�����, lngFileID, False, 0
            If mObjTabEpr.zlExportXML(strF) Then
                MsgBox "�ɹ�����ΪXML�ļ���" & vbCrLf & "�ļ���:" & strF, vbOKOnly + vbInformation, gstrSysName
            End If
        Else
            Dim DocXML As New cEPRDocument
            '��ͨסԺ����
            DocXML.InitEPRDoc cprEM_�޸�, cprET_�����ļ�����, lngFileID
            DocXML.KeepRTF = True
            DocXML.OpenEPRDoc DocXML.frmEditor.Editor1
            If DocXML.ExportToXMLFile(DocXML.frmEditor.Editor1, strF) Then
                DoEvents
            End If
        End If
    Case conMenu_File_Exit: Unload Me
    
    Case conMenu_Edit_NewItem
        If Me.rptList.FocusedRow Is Nothing Then
            lngCopyId = 0
        ElseIf Me.rptList.FocusedRow.GroupRow = True Then
            lngCopyId = 0
        Else
            lngCopyId = Me.rptList.FocusedRow.Record.Item(mCol.ID).Value
        End If
        lngFileID = frmEPRFileEdit.ShowMe(Me, mstrKinds, True, lngCopyId)
        If lngFileID <> 0 Then Call Me.zlRefList(lngFileID)
    
    Case conMenu_Edit_Modify
        If Me.rptList.FocusedRow Is Nothing Then Exit Sub
        If Me.rptList.FocusedRow.GroupRow = True Then Exit Sub
        lngFileID = Me.rptList.FocusedRow.Record.Item(mCol.ID).Value
        lngFileID = frmEPRFileEdit.ShowMe(Me, mstrKinds, False, lngFileID)
        If lngFileID <> 0 Then Call Me.zlRefList(lngFileID)
    
    Case conMenu_Edit_Delete
        With Me.rptList
            If .FocusedRow Is Nothing Then Exit Sub
            If .FocusedRow.GroupRow Then Exit Sub
            '���Ҫɾ������ר�����µ�������Ƿ񻹴�����ר�����µ��ļ�,���һ���ļ�������ɾ��
            If mintCurKind = 3 And mstrCurFixed = "����" And Me.rptList.FocusedRow.Record(mCol.����).Value = "1" Then
                If IsLastWaveFile = True Then
                    MsgBox "ר�����µ��ļ�����Ҫ����һ��,�ķ��ļ�����ɾ����", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
            If MsgBox("���ɾ�����ļ���" & vbCrLf & "����" & .FocusedRow.Record(mCol.����).Value, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                gstrSQL = "Zl_�����ļ��б�_Delete(" & .FocusedRow.Record(mCol.ID).Value & ")"
                Err = 0: On Error GoTo errHand
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                Err = 0: On Error GoTo 0
                lngCopyId = .FocusedRow.Record.Index
                Call .Records.RemoveAt(.FocusedRow.Record.Index)
                .Populate
                If .Records.Count <> 0 Then
                    If lngCopyId >= .Records.Count Then lngCopyId = 0
                    lngFileID = .Records(lngCopyId).Item(mCol.ID).Value
                Else
                    lngFileID = 0
                End If
                Call Me.zlRefList(lngFileID)
            End If
        End With
    Case conMenu_Edit_ApplyTo
        If mlngCurFileId = 0 Then Exit Sub
        If frmEPRFileApplyTo.ShowMe(Me, mlngCurFileId) Then Call mfrmRequest.zlRefresh(mlngCurFileId)
    Case conMenu_Edit_Request
        If mlngCurFileId = 0 Then Exit Sub
        Select Case mintCurKind
        Case 1, 2, 4
            If frmEPRFileTimeout.ShowMe(Me, mlngCurFileId) Then Call mfrmRequest.zlRefresh(mlngCurFileId)
        Case 5
            If frmEPRFileDisease.ShowMe(Me, mlngCurFileId) Then Call mfrmRequest.zlRefresh(mlngCurFileId)
        Case 6
            If frmEPRFileMeasure.ShowMe(Me, mlngCurFileId) Then Call mfrmRequest.zlRefresh(mlngCurFileId)
        End Select
    Case conMenu_Edit_Compend
        If mlngCurFileId = 0 Then Exit Sub
        If mintCurKind = 3 Then
            '�����¼��ʽ����
            If mstrCurFixed = "����" Then
                If Me.rptList.FocusedRow Is Nothing Then Exit Sub
                If Me.rptList.FocusedRow.GroupRow Then Exit Sub
                If Me.rptList.FocusedRow.Record(mCol.����).Value = "1" Then
                    If frmTendWaveStyle.ShowMe(Me, mlngCurFileId) = True Then
                        Me.rptList.Tag = ""
                        Call rptList_SelectionChanged
                    End If
                Else
                    Call frmTendWavePrintSet.ShowMe(Me, mlngCurFileId)
                End If
            ElseIf mstrCurFixed = "����" Then
                '����ͼ
                If frmTendPartogramStyle.ShowMe(Me, mlngCurFileId) Then
                    Me.rptList.Tag = ""
                    Call rptList_SelectionChanged
                End If
            Else
                
                If frmTendFileStyle.ShowMe(Me, mlngCurFileId) Then
                    Me.rptList.Tag = ""
                    Call rptList_SelectionChanged
                End If
            End If
            
        ElseIf mintCurKind = 2 And mstrCurFixed = "����" Then
        
        ElseIf mstrCurFixed = "���" Then
            On Error GoTo errHand
            mObjTabEpr.InitOpenEPR Me, cprEM_�޸�, cprET_�����ļ�����, mlngCurFileId
        Else
            Dim Doc As New cEPRDocument
            If mlngCurFileId = 0 Then Exit Sub
            Doc.InitEPRDoc cprEM_�޸�, cprET_�����ļ�����, mlngCurFileId
            Doc.ShowEPREditor Me
        End If
    Case conMenu_Edit_ElementChange
        frmElementChange.ShowMe Me, mlngCurFileId
    Case conMenu_Edit_Privacy
        '��˽��������
        Dim frmP As New frmPrivacyProtect
        frmP.ShowMe Me
    Case conMenu_View_ToolBar_Button
        Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each cbrControl In Me.cbsThis(2).Controls
            cbrControl.STYLE = IIf(cbrControl.STYLE = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
        Me.cbsThis.RecalcLayout
    Case conMenu_View_StatusBar
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsThis.RecalcLayout
    
    Case conMenu_View_Jump
'        If Screen.ActiveForm.Name = mfrmFileTab.Name Then
'            Call Me.dkpMan.Panes(conPane_Request).Select
'        ElseIf Screen.ActiveForm.Name = mfrmRequest.Name Then
'            Call Me.dkpMan.Panes(conPane_Compend).Select
'        Else
'            Call Me.dkpMan.Panes(conPane_FileTab).Select
'        End If
    Case conMenu_View_LocationItem
        txtFind.SetFocus
    Case conMenu_View_Refresh
        Call zlRefList(mlngCurFileId)
    
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Forum '������̳
        Call zlWebForum(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
    Case Else
        'ִ�з�������ǰģ��ı���
        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            If rptList.SelectedRows.Count > 0 Then
                If Not rptList.SelectedRows(0).GroupRow Then
                    str��� = rptList.SelectedRows(0).Record(mCol.���).Value
                    str���� = rptList.SelectedRows(0).Record(mCol.����).Value
                End If
            End If
            If str���� <> "" Then
                Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, "���=" & str���, "����=" & str����)
            Else
                Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me)
            End If
        End If
    End Select

    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnZWave As Boolean
    If Me.Visible = False Then Exit Sub
        
    If mblnFindTag = True Then
        txtFind.ForeColor = vbBlack
        If txtFind.Text = "���������ƻ�ƴ������" Then txtFind.Text = ""
    Else
        If txtFind.Text = "" Then txtFind.ForeColor = vbGrayText: txtFind.Text = "���������ƻ�ƴ������"
    End If
    
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = (mstrKinds <> "")
        End Select
    End If
    
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel: Control.Enabled = (Me.rptList.Records.Count <> 0)
    Case conMenu_File_ExportToXML
        Control.Enabled = (Me.rptList.Records.Count <> 0)
    Case conMenu_Edit_NewItem: Control.Enabled = (mstrKinds <> "" And InStr(1, mstrPrivs, "�ļ���ɾ��") > 0)
    Case conMenu_Edit_Modify: Control.Enabled = (mlngCurFileId <> 0 And InStr(1, mstrPrivs, "�ļ���ɾ��") > 0) ' And mstrCurFixed <> "����"
    Case conMenu_Edit_Delete
        'ר�����µ�����ɾ��,��׼���µ�����ɾ��
        blnZWave = False
        If Not rptList.FocusedRow Is Nothing Then
            If Not rptList.FocusedRow.GroupRow Then
                 If mstrCurFixed = "����" And mintCurKind = 3 And rptList.FocusedRow.Record(mCol.����).Value = "1" Then
                    blnZWave = True
                 End If
            End If
        End If
        Control.Enabled = (mlngCurFileId <> 0 And InStr(1, mstrPrivs, "�ļ���ɾ��") > 0) And (Trim(mstrCurFixed) = "" Or mstrCurFixed = "���" Or mstrCurFixed = "���" Or blnZWave = True)
    Case conMenu_Edit_ApplyTo: Control.Enabled = (mlngCurFileId <> 0 And Not mblnPartogram And InStr(1, mstrPrivs, "���ÿ���") > 0)
    Case conMenu_Edit_Request: Control.Enabled = (mlngCurFileId <> 0 And InStr(1, mstrPrivs, "����Ҫ��") > 0) And mintCurKind <> 3
    Case conMenu_Edit_Compend
        Control.Enabled = (mlngCurFileId <> 0 And InStr(1, mstrPrivs, "��ʽ����") > 0)
        If Control.Enabled Then Control.Enabled = (mintCurKind <> 3 Or mintCurKind = 3 And mstrCurFixed <> "����")
        If Control.Enabled Then Control.Enabled = mstrCurFixed <> "����"
    Case conMenu_Edit_Privacy: Control.Enabled = (InStr(1, mstrPrivs, "��˽����") > 0)
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).STYLE = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    Case conMenu_Edit_ElementChange: Control.Enabled = (mlngCurFileId <> 0) And Not (mstrCurFixed = "���" Or mstrCurFixed = "���" Or mstrCurFixed = "����" Or mintCurKind = 3)
    End Select
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_FileTab
        Item.Handle = Me.PicFileTab.hWnd
    Case conPane_Request
        If mfrmRequest Is Nothing Then Set mfrmRequest = New frmEPRFileRequest
        Item.Handle = mfrmRequest.hWnd
    Case conPane_Compend
        If mfrmContent Is Nothing Then Set mfrmContent = New frmEPRFileContent
        Item.Handle = mfrmContent.hWnd
    End Select
End Sub

Private Sub Form_Load()
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar
Dim rptCol As ReportColumn
Dim lngCount As Long

    '-----------------------------------------------------
    'Ȩ�����ƴ����ƣ�����ͬʱ��������ģ�������gstrPrivs�仯�����¿�����Ч
    mstrPrivs = gstrPrivs
    mstrKinds = ""
    mblnPartogram = False
    If InStr(1, mstrPrivs, "���ﲡ��") > 0 Then mstrKinds = mstrKinds & ",1"
    If InStr(1, mstrPrivs, "סԺ����") > 0 Then mstrKinds = mstrKinds & ",2"
    If InStr(1, mstrPrivs, "�����¼") > 0 Then mstrKinds = mstrKinds & ",3"
    If InStr(1, mstrPrivs, "������") > 0 Then mstrKinds = mstrKinds & ",4"
    If InStr(1, mstrPrivs, "����֤������") > 0 Then mstrKinds = mstrKinds & ",5"
    If InStr(1, mstrPrivs, "֪���ļ�") > 0 Then mstrKinds = mstrKinds & ",6"
    If mstrKinds <> "" Then mstrKinds = Mid(mstrKinds, 2)
    
    Call ZLCommFun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
    
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set Me.cbsThis.Icons = ZLCommFun.GetPubIcons
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
        Set cbrControl = .Add(xtpControlButton, conMenu_File_ExportToXML, "����ΪXML�ļ�(&L)��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_ExportToXML + 1, "��������XML�ļ�(&E)��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_ExportToXML + 2, "��������XML�ļ�(&I)��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ApplyTo, "���ÿ���(&T)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Request, "����Ҫ��(&R)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Compend, "��ʽ����(&F)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ElementChange, "Ҫ����������(&E)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Privacy, "��˽��Ŀ����(&P)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_View_LocationItem, "����(&F)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Jump, "������ת(&J)")
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "��̳(&F)", -1, False  '����
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): cbrControl.BeginGroup = True
    End With
    
    '�����
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add 0, VK_DELETE, conMenu_Edit_Delete
        .Add FCONTROL, Asc("T"), conMenu_Edit_ApplyTo
        .Add FCONTROL, Asc("R"), conMenu_Edit_Request
        .Add FCONTROL, Asc("D"), conMenu_Edit_Compend
        .Add FCONTROL, Asc("F"), conMenu_View_LocationItem
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F6, conMenu_View_Jump
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '���ò����ò˵�
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Jump
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
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ApplyTo, "ʹ�ÿ���"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Request, "����Ҫ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Compend, "��ʽ����")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.STYLE = xtpButtonIconAndCaption
    Next
    
    '��ȡ��������ģ��ı���:��Ϊ��һ���Զ�ȡ,ȫ�ֱ�������
    '---------------------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
    
    '-----------------------------------------------------
    '���ôʾ���ʾͣ������
    If mfrmRequest Is Nothing Then Set mfrmRequest = New frmEPRFileRequest
    If mfrmContent Is Nothing Then Set mfrmContent = New frmEPRFileContent
    If mObjTabEpr Is Nothing Then Set mObjTabEpr = New cTableEPR
    mObjTabEpr.InitTableEPR gcnOracle, glngSys, gstrDbOwner
    
    Dim panFileTab As Pane, panRequest As Pane, panCompend As Pane
    Set panFileTab = dkpMan.CreatePane(conPane_FileTab, 180, 400, DockLeftOf, Nothing)
    panFileTab.Title = "�ļ��б�"
    panFileTab.Options = PaneNoCaption
    
    Set panRequest = dkpMan.CreatePane(conPane_Request, 400, 200, DockRightOf, Nothing)
    panRequest.Title = "Ӧ��Ҫ��"
    panRequest.Options = PaneNoCloseable Or PaneNoFloatable 'Or PaneNoHideable

    Set panCompend = dkpMan.CreatePane(conPane_Compend, 400, 300, DockBottomOf, panRequest)
    panCompend.Title = "�ļ���ʽ"
    panCompend.Options = PaneNoCloseable Or PaneNoFloatable 'Or PaneNoHideable
    
    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
    
    '-----------------------------------------------------
    With Me.rptList
        Set rptCol = .Columns.Add(mCol.ͼ��, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mCol.ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.����, "����", 90, False): rptCol.Editable = False: rptCol.Groupable = True: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.���, "���", 50, False): rptCol.Editable = False: rptCol.Groupable = False: .SortOrder.Add rptCol
        Set rptCol = .Columns.Add(mCol.����, "����", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.˵��, "˵��", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.����, "����", 30, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.ҳ��, "ҳ��", 80, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.����, "����", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.����, "����", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        
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
    '��ѯ���ʼ��
    mblnFindTag = False
    txtFind.ForeColor = vbGrayText
    txtFind.Text = "���������ƻ�ƴ������"
    mintLastRows = 0
    
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
        Me.stbThis.Panels(2).Text = "����" & lngCount & "�������ļ�"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload mfrmRequest
    Unload mfrmContent
    Set mfrmRequest = Nothing
    Set mfrmContent = Nothing
    Set mObjTabEpr = Nothing
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub mfrmContent_DblClick()
    Dim cbrControl As CommandBarControl
    If mlngCurFileId = 0 Or mstrCurFixed = "����" Then Exit Sub
    Set cbrControl = Me.cbsThis.FindControl(, conMenu_Edit_Compend)
    If cbrControl Is Nothing Then Exit Sub
    If cbrControl.Visible = False Or cbrControl.Enabled = False Then Exit Sub
    Call cbsThis_Execute(cbrControl)
End Sub
Private Sub mfrmRequest_DblClick(lngWhere As zlEnumDClick)
Dim cbrControl As CommandBarControl
    If mlngCurFileId = 0 Then Exit Sub
    Select Case lngWhere
    Case cprEmDClickApplyTo: Set cbrControl = Me.cbsThis.FindControl(, conMenu_Edit_ApplyTo)
    Case cprEmDClickRequest: Set cbrControl = Me.cbsThis.FindControl(, conMenu_Edit_Request)
    Case Else: Set cbrControl = Nothing
    End Select
    If cbrControl Is Nothing Then Exit Sub
    If cbrControl.Visible = False Or cbrControl.Enabled = False Then Exit Sub
    Call cbsThis_Execute(cbrControl)
End Sub


Private Sub PicFileTab_Resize()
    lblFind.Move 70, 90, lblFind.Width, lblFind.Height
    If PicFileTab.Width > 800 Then txtFind.Move 800, 50, PicFileTab.Width - 800, 300
    If PicFileTab.Height > 400 Then rptList.Move 0, 400, PicFileTab.Width, PicFileTab.Height - 400
End Sub

Private Sub rptList_KeyDown(KeyCode As Integer, Shift As Integer)
    If Me.rptList.Visible = False Then Exit Sub
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Me.rptList.FocusedRow Is Nothing Then Exit Sub
    If Me.rptList.FocusedRow.GroupRow Then Exit Sub
    Call rptList_RowDblClick(Me.rptList.FocusedRow, Me.rptList.FocusedRow.Record.Item(mCol.���))
End Sub

Private Sub rptList_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
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

Private Sub rptList_RowDblClick(ByVal ROW As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
Dim cbrControl As CommandBarControl

    With Me.rptList
        If .FocusedRow Is Nothing Then
            mintCurKind = 0: mlngCurFileId = 0: mstrCurFixed = "": mblnPartogram = False
        ElseIf .FocusedRow.GroupRow = True Then
            mintCurKind = .FocusedRow.Childs.ROW(0).Record.Item(mCol.ͼ��).Value: mlngCurFileId = 0: mstrCurFixed = "": mblnPartogram = False
        Else
            mintCurKind = .FocusedRow.Record.Item(mCol.ͼ��).Value
            mlngCurFileId = .FocusedRow.Record.Item(mCol.ID).Value
            mstrCurFixed = .FocusedRow.Record.Item(mCol.����).Value
            mblnPartogram = ((mstrCurFixed = "����") And (mintCurKind = 3))
        End If
    End With
    If mlngCurFileId = 0 Then Exit Sub
    
    Set cbrControl = Me.cbsThis.FindControl(, conMenu_Edit_Modify)
    If cbrControl Is Nothing Then Exit Sub
    If cbrControl.Visible = False Or cbrControl.Enabled = False Then Exit Sub
    Call cbsThis_Execute(cbrControl)

End Sub

Private Sub rptList_SelectionChanged()
    With Me.rptList
        If .FocusedRow Is Nothing Then
            mintCurKind = 0: mlngCurFileId = 0: mstrCurFixed = "": mblnPartogram = False
        ElseIf .FocusedRow.GroupRow = True Then
            mintCurKind = .FocusedRow.Childs.ROW(0).Record.Item(mCol.ͼ��).Value: mlngCurFileId = 0: mstrCurFixed = "": mblnPartogram = False
        Else
            mintCurKind = .FocusedRow.Record.Item(mCol.ͼ��).Value
            mlngCurFileId = .FocusedRow.Record.Item(mCol.ID).Value
            mstrCurFixed = .FocusedRow.Record.Item(mCol.����).Value
            mblnPartogram = ((mstrCurFixed = "����") And (mintCurKind = 3))
            If Val(Me.rptList.Tag) <> Me.rptList.FocusedRow.Index Then
                Call mfrmRequest.zlRefresh(mlngCurFileId)
                Call mfrmContent.zlRefresh(mlngCurFileId)
                Me.rptList.Tag = Me.rptList.FocusedRow.Index
            End If
        End If
    End With
End Sub

Private Function IsLastWaveFile() As Boolean
'����:���ר�����µ��ļ��Ƿ������һ��
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    gstrSQL = "Select Count(1) ��Ŀ From �����ļ��б� where ����=3 And ����=-1 And ����='1'"
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "�����ļ��б�")
    IsLastWaveFile = rsTemp!��Ŀ = 1
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub txtFind_Change()
    mintLastRows = 0
End Sub

Private Sub txtFind_GotFocus()
    mblnFindTag = True
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack Then txtFind.SetFocus '��ֹ��ɾ��������ת��bug
End Sub

Private Sub txtFind_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim intCount As Integer

    If KeyCode = vbKeyReturn And txtFind.Text <> "" Then
        For intCount = mintLastRows + 1 To Me.rptList.Rows.Count - 1
            If Me.rptList.Rows(intCount).GroupRow = False Then
                If InStr(Me.rptList.Rows(intCount).Record(mCol.����).Value, txtFind.Text) Or InStr(Me.rptList.Rows(intCount).Record(mCol.����).Value, UCase(txtFind.Text)) Then
                    Set Me.rptList.FocusedRow = Me.rptList.Rows(intCount)
                    mintLastRows = intCount
                    Exit For
                End If
            End If
        Next
        If intCount = Me.rptList.Rows.Count And mintLastRows = 0 Then
            Call MsgBox("δ�ҵ��롰" & txtFind.Text & "��ƥ��Ĳ������������������ƻ���롣", vbInformation, gstrSysName)
            txtFind.Text = ""
        End If
    End If
    txtFind.SetFocus
End Sub

Private Sub txtFind_LostFocus()
    mblnFindTag = False
End Sub
