VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{FBAFE9A8-8B26-4559-9D12-D70E36A97BE3}#2.1#0"; "zlRichEditor.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFileExportOrImport 
   Caption         =   "�����ļ������б�"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8775
   Icon            =   "frmFileExportOrImport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   8775
   StartUpPosition =   1  '����������
   Begin zlRichEditor.Editor Editor1 
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   2400
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   3960
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSFlex8Ctl.VSFlexGrid vsgrid 
      Height          =   1860
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   2655
      _cx             =   4683
      _cy             =   3281
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
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
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
   Begin VB.PictureBox PicBtn 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   8775
      TabIndex        =   1
      Top             =   5160
      Width           =   8775
      Begin MSComctlLib.ProgressBar progBar 
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   5520
      Top             =   1080
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
            Picture         =   "frmFileExportOrImport.frx":6852
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileExportOrImport.frx":6DEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileExportOrImport.frx":7386
            Key             =   "ǩ��"
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox RTbFootText 
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      _Version        =   393217
      TextRTF         =   $"frmFileExportOrImport.frx":76D8
   End
   Begin RichTextLib.RichTextBox RTbHeadText 
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      _Version        =   393217
      TextRTF         =   $"frmFileExportOrImport.frx":7775
   End
   Begin RichTextLib.RichTextBox RTbContext 
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   3360
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      _Version        =   393217
      TextRTF         =   $"frmFileExportOrImport.frx":7812
   End
   Begin XtremeCommandBars.ImageManager imgManager 
      Left            =   6360
      Top             =   1200
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmFileExportOrImport.frx":78AF
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
End
Attribute VB_Name = "frmFileExportOrImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Private mdoc As DOMDocument         'Xml�ĵ�
Private mIntClosed As Integer       '���ƴ����Ƿ���Թر�
Private mstrPath As String          '·��
Private mblnInit As Boolean         '������Enable״̬
Private mlngType As Long            '��ǰ���崦�ڵ���/����״̬��1Ϊ������2Ϊ���룩
Private Type mDocType
    mDocXML As New DOMDocument
    mXmlPath As String
End Type
Private Enum mExportCols
    Range = 0: Choose: cType: cID: cNull: cName
End Enum
Private Enum mImportCols
    Choose = 0: cName: cImportType: cTip: cUnit: cPath
End Enum
Private mDocArr() As mDocType
Private Enum menu_this
    menu_Cover = 1                  '���븲���ļ�
    menu_Add = 2                    '���������ļ�
    menu_RemoveRow = 3              '�Ƴ�ѡ����
    menu_Clear = 4                  '����б�
    menu_Export = 10                '����
    menu_Import = 11                '����
    menu_IcheckAll = 12             '����ȫѡ
    menu_IclearAll = 13             '����ȫ��
    menu_CheckThis = 14             'ѡ��ǰ��
    menu_AddFile = 15               '���������ļ�
    menu_Unload = 16                '�˳�
    menu_EcheckAll = 17             '����ȫѡ
    menu_EclearAll = 18             '����ȫ��
    menu_ExportOne = 101            '����һ��XML
    menu_ExportMore = 102           '�������XML
    menu_CheckHave = 121            'ȫѡ�Ѵ����ļ�
    menu_ClearHave = 131            'ȫ���Ѵ����ļ�
    menu_ImportOption = 110         '��������
End Enum
Public Function ShowMe(ByVal objParent As Object, ByVal lngType As Long)
    If lngType = 1 Then
        Call ExportList
        Me.Caption = "�����ļ������б�"
        Me.vsgrid.Tag = "Export"
        Me.Tag = ""
        Me.ShowControl(Me.cbsThis, 11, True).Visible = False
        Me.Show 1, objParent
    Else
        If ImportList Then
        Me.Caption = "�����ļ������б�"
        Me.vsgrid.Tag = "Import"
        Me.ShowControl(Me.cbsThis, 10, True).Visible = False
        Me.Show 1, objParent
        End If
    End If
End Function
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case menu_Cover
             If vsgrid.Row < 1 Then Exit Sub
             vsgrid.Cell(flexcpForeColor, vsgrid.Row, 0, vsgrid.Row, 3) = vbMagenta
             vsgrid.TextMatrix(vsgrid.Row, 3) = "�Ѵ���,���뽫����ԭ���ļ���"
             vsgrid.TextMatrix(vsgrid.Row, 2) = Split(vsgrid.TextMatrix(vsgrid.Row, 2), "_")(0) & "_1"
             vsgrid.Cell(flexcpData, vsgrid.Row, 3) = "����"
             vsgrid_DblClick
        Case menu_Add
             If vsgrid.Row < 1 Then Exit Sub
             vsgrid.Cell(flexcpForeColor, vsgrid.Row, 0, vsgrid.Row, 3) = vbBlack
             vsgrid.TextMatrix(vsgrid.Row, 3) = "�Ѵ���,���뽫�������ļ���"
             vsgrid.TextMatrix(vsgrid.Row, 2) = Split(vsgrid.TextMatrix(vsgrid.Row, 2), "_")(0) & "_2"
             vsgrid.Cell(flexcpData, vsgrid.Row, 3) = ""
             If vsgrid.Cell(flexcpPicture, vsgrid.Row, 0) Is Nothing Then vsgrid_DblClick
        Case menu_RemoveRow
             vsgrid.RemoveItem (vsgrid.Row)
             If vsgrid.Rows = 1 Then InitVsGrid ("���������Ҫ����Ĳ����ļ� ��")
             Me.Tag = Val(Me.Tag) - 1
        Case menu_Clear
            Call InitVsGrid("���������Ҫ����Ĳ����ļ� ��")
        Case menu_Export
            Call Export
        Case menu_ExportOne
             If vsgrid.Cols > 1 Then
                Control.Checked = True
                ShowControl(Me.cbsThis, 102, True).Checked = False
             End If
        Case menu_ExportMore
             If vsgrid.Cols > 1 Then
                Control.Checked = True
                ShowControl(Me.cbsThis, 101, True).Checked = False
             End If
        Case menu_Import
            Call Import
        Case menu_IcheckAll, menu_EcheckAll
            Call CheckItems(True)
        Case menu_CheckHave
            Call CheckHave(True)
        Case menu_ClearHave
            Call CheckHave(False)
        Case menu_IclearAll, menu_EclearAll
            Call CheckItems(False)
        Case menu_CheckThis
            Call vsgridClick
        Case menu_AddFile
            Call ImportList
            Me.vsgrid.Tag = "Import"
        Case menu_Unload
            Unload Me
         Exit Sub
    End Select
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If mblnInit Then
        Control.Enabled = False
    Else
       Select Case Control.ID
           Case menu_Cover
                Control.Enabled = IIf(vsgrid.TextMatrix(vsgrid.Row, 3) = "", False, True)
                Control.Checked = IIf(vsgrid.TextMatrix(vsgrid.Row, 3) = "�Ѵ���,���뽫����ԭ���ļ���", True, False)
           Case menu_Add
                Control.Enabled = IIf(vsgrid.TextMatrix(vsgrid.Row, 3) = "", False, True)
                Control.Checked = IIf(vsgrid.TextMatrix(vsgrid.Row, 3) = "�Ѵ���,���뽫�������ļ���", True, False)
           Case menu_Export, menu_Import
                Control.Enabled = IIf(vsgrid.Tag = "" Or vsgrid.Cols = 1, False, True)
           Case menu_ExportOne, menu_ExportMore
                Control.Enabled = IIf(vsgrid.Tag = "Import", False, True)
           Case menu_IcheckAll, menu_IclearAll, menu_ImportOption
                Control.Enabled = IIf(vsgrid.Rows = 1, False, True)
                Control.Visible = IIf(vsgrid.Tag = "Export", False, True)
           Case menu_CheckThis
                Control.Visible = False
           Case menu_AddFile
                Control.Visible = IIf(vsgrid.Tag = "Export", False, True)
           Case menu_EcheckAll, menu_EclearAll
                Control.Visible = IIf(vsgrid.Tag = "Import", False, True)
                Control.Enabled = IIf(vsgrid.Rows = 1, False, True)
        End Select
    End If
End Sub
Private Sub CheckItems(ByVal blnOn As Boolean)
    Dim i As Integer, intCol As Integer, j As Integer
    intCol = IIf(vsgrid.Tag = "Import", 0, 1)
    For i = 1 To vsgrid.Rows - 1
         If intCol = 0 Then
            vsgrid.Cell(flexcpPicture, i, 0) = IIf(blnOn And Not IsHave(i), img16.ListImages("Check").Picture, Nothing)
            vsgrid.Cell(flexcpData, i, 0) = IIf(blnOn And Not IsHave(i), 1, 0)
         Else
            vsgrid.Cell(flexcpPicture, i, 1) = IIf(blnOn, img16.ListImages("Check").Picture, Nothing)
            vsgrid.Cell(flexcpData, i, 1) = IIf(blnOn, 1, 0)
         End If
    Next i
End Sub
Private Sub Form_Load()
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim blnChecked As Boolean
    Set objBar = Me.cbsThis.Add("Tools", xtpBarTop)
    objBar.ContextMenuPresent = False           '�������ϵ������Ҽ�ʱ���������ò˵�
    objBar.ShowTextBelowIcons = False           '�������еİ�ť������ʾ��ͼ���Ҳ�
    objBar.EnableDocking xtpFlagStretched
    Me.cbsThis.Icons = Me.imgManager.Icons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True                 '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    With objBar.Controls
        Set objPopup = .Add(xtpControlSplitButtonPopup, 10, "����"): objPopup.STYLE = xtpButtonIconAndCaption
            objPopup.BeginGroup = True
            objPopup.ID = 10                    'Popup��ID�����¸�ֵ������Ч
            objPopup.IconId = 10                'Popup��IconId�����¸�ֵ������Ч
        objPopup.CommandBar.Width = 100
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, 101, "����Ϊһ��XML�ļ�(&Q)"
            .Add xtpControlButton, 102, "����Ϊ���XML�ļ�(&W)"
        End With
        Set objControl = .Add(xtpControlButton, 11, "����"): objControl.STYLE = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, 15, "���"): objControl.STYLE = xtpButtonIconAndCaption
        Set objPopup = .Add(xtpControlButtonPopup, 110, "����ѡ��"): objPopup.STYLE = xtpButtonIconAndCaption
            objPopup.BeginGroup = True
            objPopup.ID = 110
            objPopup.IconId = 110
        objPopup.CommandBar.Width = 60
        With objPopup.CommandBar.Controls
        .Add xtpControlButton, 1, "���뽫����ԭ���ļ�"
        .Add xtpControlButton, 2, "���뽫�������ļ�"
        End With
        Set objControl = .Add(xtpControlButton, 17, "ȫѡ"): objControl.STYLE = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, 18, "ȫ��"): objControl.STYLE = xtpButtonIconAndCaption
        Set objPopup = .Add(xtpControlSplitButtonPopup, 12, "ȫѡ"): objPopup.STYLE = xtpButtonIconAndCaption
            objPopup.BeginGroup = True
            objPopup.ID = 12
            objPopup.IconId = 12
        objPopup.CommandBar.Width = 60
        With objPopup.CommandBar.Controls
        .Add xtpControlButton, 121, "ȫѡ�Ѵ����ļ�"
        End With
        Set objPopup = .Add(xtpControlSplitButtonPopup, 13, "ȫ��"): objPopup.STYLE = xtpButtonIconAndCaption
            objPopup.BeginGroup = True
            objPopup.ID = 13
            objPopup.IconId = 13
        objPopup.CommandBar.Width = 60
        With objPopup.CommandBar.Controls
        .Add xtpControlButton, 131, "ȫ���Ѵ����ļ�"
        End With
        Set objControl = .Add(xtpControlButton, 14, "ѡ��"): objControl.STYLE = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, 16, "�˳�"): objControl.STYLE = xtpButtonIconAndCaption
        objControl.STYLE = xtpButtonIconAndCaption  'ͬʱ��ʾͼ�������
    End With
    blnChecked = IIf(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ExportType", App.Path) = "More", False, True)
    Me.cbsThis.ActiveMenuBar.Visible = False
    Me.ShowControl(Me.cbsThis, 101, False).Checked = blnChecked
    Me.ShowControl(Me.cbsThis, 102, False).Checked = Not blnChecked
End Sub
Private Sub Export()
    Dim i As Integer, lngArrFile As Variant, strItems As String
    '�õ�����ѡ����
     For i = 1 To vsgrid.Rows - 1
         If Not vsgrid.Cell(flexcpPicture, i, mExportCols.Choose) Is Nothing And vsgrid.GetNode(i).Children < 1 Then
             strItems = strItems & vsgrid.TextMatrix(i, mExportCols.cID) & "_" & vsgrid.TextMatrix(i, mExportCols.cName) & "_" & vsgrid.TextMatrix(i, mExportCols.cType) & ","
         End If
     Next i
     If strItems = "" Then
         MsgBox "��ѡ����Ҫ�������ļ���", vbInformation, gstrSysName
         Exit Sub
     End If
     strItems = Mid(strItems, 1, Len(strItems) - 1)
     'ָ��������ļ�·��
     mstrPath = zl9ComLib.OS.OpenDir(Me.hWnd, "ָ������Ŀ¼")
     If mstrPath = "" Then Exit Sub
     On Error Resume Next
     mstrPath = mstrPath & "\" & zl9ComLib.GetUnitName
     gobjFSO.CreateFolder (mstrPath)
     If Err.Number = 32755 Then Err.Clear: Exit Sub
     On Error GoTo errHand
     lngArrFile = Split(strItems, ",")
     mIntClosed = 1: EnableControlBar Me, False: mblnInit = True
     Call StartExportToXMLFile(cprEM_�޸�, cprET_�����ļ�����, lngArrFile, IIf(ShowControl(Me.cbsThis, 101, True).Checked, 1, 2))
     MsgBox "�����ѵ�����Ŀ���ļ���" & mstrPath, vbApplicationModal + vbInformation, "����"
     mIntClosed = 0: mblnInit = False
     Unload Me
     Exit Sub
errHand:
    mIntClosed = 0: mblnInit = False
    EnableControlBar Me, True
    If ErrCenter() = 1 Then
         Resume
    End If
    Call SaveErrLog
End Sub
Private Sub Import()
    Dim i As Long, j As Long
    On Error GoTo errHand
    '��ʼѭ������
    '---------------
    Me.progBar.Visible = True
    mIntClosed = 1: EnableControlBar Me, False: mblnInit = True
    ReDim Preserve mDocArr(1 To 1) As mDocType
    For i = 1 To vsgrid.Rows - 1
        Set mdoc = Nothing
        If Not vsgrid.Cell(flexcpPicture, i, mImportCols.Choose) Is Nothing Then
            If UBound(mDocArr) > 1 Then
                For j = 1 To UBound(mDocArr)
                    If vsgrid.TextMatrix(i, mImportCols.cPath) = mDocArr(j).mXmlPath Then Set mdoc = mDocArr(j).mDocXML: Exit For
                Next j
            End If
            If mdoc Is Nothing Then
                ReDim Preserve mDocArr(1 To UBound(mDocArr) + 1) As mDocType
                mDocArr(UBound(mDocArr)).mDocXML.Load vsgrid.TextMatrix(i, mImportCols.cPath)
                Set mdoc = mDocArr(UBound(mDocArr)).mDocXML
                mDocArr(UBound(mDocArr)).mXmlPath = vsgrid.TextMatrix(i, mImportCols.cPath)
            End If
            DoEvents
            Me.Refresh
            Call ImportFromXml(vsgrid.TextMatrix(i, mImportCols.cPath), vsgrid.TextMatrix(i, mImportCols.cImportType))
        End If
        progBar.Value = IIf(progBar.Value + progBar.Max / (vsgrid.Rows - 1) > progBar.Max, progBar.Max, progBar.Value + progBar.Max / (vsgrid.Rows - 1))
    Next i
    '----------------
    MsgBox "������� ��", vbInformation, gstrSysName
    mIntClosed = 0: mblnInit = False
    Unload Me
    Exit Sub
errHand:
    mIntClosed = 0: mblnInit = False
    EnableControlBar Me, True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
'���ز����ļ��б�
Public Sub ExportList()
    Dim strType As String, i As Long, j As Long, k As Long
    Dim rsTemp As ADODB.Recordset
    gstrSQL = "select distinct ID,����,���,���� from (" & _
              "  Select l.Id, decode(l.����,1,'1-���ﲡ��',2,'2-סԺ����',4,'4-������',5,'5-����֤������',6,'6-֪���ļ�') as ����, l.���, l.����" & _
              "  From �����ļ��б� l where l.����<>2 and l.���� in (1,2,4,5,6)" & _
              "  Union All Select 0, decode(l.����,1,'1-���ﲡ��',2,'2-סԺ����',4,'4-������',5,'5-����֤������',6,'6-֪���ļ�') as ����, null, null" & _
              "  From �����ļ��б� l where l.����<>2 and l.���� in (1,2,4,5,6)) order by ����,ID"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        'vsgridClick
        With vsgrid
            '��������
            If rsTemp.RecordCount < 1 Then
                 Call InitVsGrid("��ʱû�п��Ե����Ĳ����ļ� ��"):  rsTemp.Close
                 Exit Sub
            End If
            .Clear: .Cols = 6: .Rows = 1: .FixedRows = 1: .ROWHEIGHT(0) = 10
            .ColWidth(mExportCols.Range) = 400: .ColWidth(mExportCols.Choose) = 270: .ColWidth(mExportCols.cType) = 0: .ColWidth(mExportCols.cNull) = 0: .ColWidth(mExportCols.cID) = 0: .ColWidth(mExportCols.cName) = 2500
            .Cell(flexcpData, 0, mExportCols.Choose) = 1: .TextMatrix(0, mExportCols.cName) = "����": .TextMatrix(0, mExportCols.Choose) = "ѡ��": .ColAlignment(mExportCols.cName) = flexAlignLeftCenter
            '���÷���
            .OutlineCol = 0: .OutlineBar = flexOutlineBarCompleteLeaf
            For i = 1 To rsTemp.RecordCount
                If strType <> rsTemp!���� Then
                    .AddItem ""
                    For k = 2 To .Cols - 1
                        .TextMatrix(i, k) = NVL(rsTemp("����").Value)
                    Next k
                    .Cell(flexcpBackColor, i, 0, i, 5) = &HFFC0C0
                    .IsSubtotal(i) = True
                    .Cell(flexcpData, i, mExportCols.Choose) = 1
                    .MergeCells = flexMergeFree
                    .MergeRow(1) = True '�Ƿ������кϲ�
                    strType = rsTemp!����
                Else
                    .AddItem ""
                    .Cell(flexcpData, i, mExportCols.Choose) = 0
                    .TextMatrix(i, mExportCols.cType) = NVL(rsTemp("����").Value)
                    .TextMatrix(i, mExportCols.cID) = NVL(rsTemp("ID").Value)
                    .TextMatrix(i, mExportCols.cName) = NVL(rsTemp("����").Value)
                    .IsSubtotal(i) = True
                    .RowOutlineLevel(i) = 1
                End If
                .Cell(flexcpPicture, i, mExportCols.Choose) = img16.ListImages("Check").Picture
                .Cell(flexcpData, i, mExportCols.Choose) = 1
                rsTemp.MoveNext
           Next i
           For i = 1 To vsgrid.Rows - 1
                If .IsSubtotal(i) = True Then
                    .GetNode(i).Expanded = True
                End If
           Next i
           If vsgrid.Rows > 1 Then vsgrid.Row = 2
    End With
    rsTemp.Close
End Sub
'ѡ���Ѵ����ļ�
Private Sub CheckHave(ByVal blnOn As Boolean)
    Dim i As Integer
    For i = 1 To vsgrid.Rows - 1
       If vsgrid.Cell(flexcpForeColor, i, 1, i, 3) = vbMagenta Then
          vsgrid.Cell(flexcpPicture, i, 0) = IIf(blnOn And Not IsHave(i), img16.ListImages("Check").Picture, Nothing)
       End If
    Next i
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Cancel = mIntClosed
End Sub

Private Sub Form_Resize()
    Me.vsgrid.Move 0, 500, Me.ScaleWidth, Me.ScaleHeight - Me.PicBtn.Height - 500
    Me.PicBtn.Move 0, vsgrid.Height + 500, Me.ScaleWidth, PicBtn.Height
    Me.progBar.Move 0, 60, PicBtn.Width, progBar.Height
    If vsgrid.Rows = 1 Or vsgrid.Cols = 1 Then
        vsgrid.ROWHEIGHT(0) = Me.ScaleHeight
    End If
End Sub

'����VsGrid
Private Sub vsgridClick()
    Dim i As Long, j As Long, strItems As String
    Dim intX As Integer, intW As Integer
    If vsgrid.Cols = 1 Then Exit Sub
    If vsgrid.Tag = "Import" Then
       If vsgrid.Row > 0 Then
         If IsHave(vsgrid.Row) And vsgrid.Cell(flexcpData, vsgrid.Row, 0) = 0 And vsgrid.Cell(flexcpData, vsgrid.Row, 3) = "����" Then
            MsgBox "����ͬʱѡ������������ͬ�Ĳ����ļ�����ԭ���ļ���", vbInformation, gstrSysName: Exit Sub
         End If
         vsgrid.Cell(flexcpData, vsgrid.Row, 0) = IIf(vsgrid.Cell(flexcpData, vsgrid.Row, 0) = 1, 0, 1)
         vsgrid.Cell(flexcpPicture, vsgrid.Row, 0) = IIf(vsgrid.Cell(flexcpData, vsgrid.Row, 0) = 1, img16.ListImages("Check").Picture, Nothing)
       End If
    Else
        'ѡ�нڵ�����������
        If vsgrid.MouseRow < 0 Then Exit Sub
        vsgrid.Cell(flexcpData, vsgrid.Row, 1) = IIf(vsgrid.Cell(flexcpData, vsgrid.Row, 1) = 1, 0, 1)
        For i = vsgrid.Row To vsgrid.GetNode(vsgrid.Row).Children + vsgrid.Row
             vsgrid.Cell(flexcpData, vsgrid.Row, 0) = IIf(vsgrid.Cell(flexcpData, vsgrid.Row, 0) = 1, 0, 1)
             vsgrid.Cell(flexcpPicture, i, 1) = IIf(vsgrid.Cell(flexcpData, vsgrid.Row, 1) = 1, img16.ListImages("Check").Picture, Nothing)
        Next i
    End If
End Sub
'
Private Function IsHave(ByVal intRow As Long) As Boolean
    Dim i As Long
    For i = 1 To vsgrid.Rows - 1
        If intRow <> i And vsgrid.Cell(flexcpData, i, 3) = "����" And Not vsgrid.Cell(flexcpPicture, i, 0) Is Nothing Then
            If vsgrid.TextMatrix(intRow, 1) = vsgrid.TextMatrix(i, 1) And vsgrid.Cell(flexcpForeColor, intRow, 1, intRow, 3) = vbMagenta Then
            IsHave = True: Exit Function
            End If
        End If
    Next i
    IsHave = False
End Function

Private Sub Form_Unload(Cancel As Integer)
  Dim ExportType As String
  ExportType = IIf(ShowControl(Me.cbsThis, 101, True).Checked, "One", "More")
  SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ExportType", ExportType
  If Not mdoc Is Nothing Then Set mdoc = Nothing
End Sub

Private Sub vsgrid_Click()
    If vsgrid.Cols = 1 Then Exit Sub
    If vsgrid.MouseIcon Is Nothing Then Exit Sub
    If vsgrid.MouseIcon = Me.img16.ListImages(1).Picture Then
            vsgridClick
    End If
End Sub

'˫���¼�
Private Sub vsgrid_DblClick()
     If vsgrid.Cols = 1 Then Exit Sub
     If vsgrid.MouseIcon Is Nothing And vsgrid.MouseRow > 1 Then
        If vsgrid.Tag = "Export" Then
            If vsgrid.GetNode(vsgrid.Row).Children > 1 Then
                vsgrid.GetNode(vsgrid.Row).Expanded = Not vsgrid.GetNode(vsgrid.Row).Expanded: Exit Sub
            End If
        End If
        vsgridClick
        Exit Sub
     End If
End Sub
'���¼����¼�
Private Sub vsgrid_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsgrid
        If .IsSubtotal(.Row) Then
            Select Case KeyCode
              Case vbKeyLeft
                  .GetNode(.Row).Expanded = False
              Case vbKeySpace
                   vsgridClick
              Case vbKeyRight
                  .GetNode(.Row).Expanded = True
              Case 13
                .GetNode(.Row).Expanded = Not .GetNode(.Row).Expanded
              Case vbKeyA
                If Shift = 2 Then CheckItems (True)
              Case vbKeyZ
                If Shift = 2 Then CheckItems (False)
            End Select
        ElseIf vsgrid.Tag = "Import" Then
            If KeyCode = 13 Then
              Call vsgridClick
            ElseIf KeyCode = vbKeySpace Then
                vsgridClick
            ElseIf KeyCode = vbKeyA Then
              If Shift = 2 Then CheckItems (True)
            ElseIf KeyCode = vbKeyZ Then
              If Shift = 2 Then CheckItems (False)
            End If
        End If
   End With
End Sub

Private Sub vsgrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngH As Long, lngY As Long
     If vsgrid.Cols = 1 Or vsgrid.Rows = 1 Then Exit Sub
     lngH = vsgrid.Row * 255: lngY = 255 * (vsgrid.Row + 1)
     If Button = 2 Then
        If vsgrid.Tag = "Import" And Y > lngH And Y < lngY Then
                Dim Popup As CommandBar
                Dim objControl As CommandBarControl
                Set Popup = cbsThis.Add("Popup", xtpBarPopup)
                With Popup.Controls
                    .Add xtpControlButton, 1, "����ʱ���Ǹò���(&F)"
                    .Add xtpControlButton, 2, "����ʱ�����ò���(&A)"
                    .Add xtpControlButton, 3, "���б����Ƴ�(&D)"
                    .Add xtpControlButton, 4, "����б�(&C)"
                End With
                Popup.ShowPopup
        End If
      End If
End Sub

Private Sub vsgrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intX As Integer, intW As Integer
    If vsgrid.Cols = 1 Then Exit Sub
    If vsgrid.Tag = "Import" Then intX = 0: intW = CSng(vsgrid.ColWidth(0)) Else intX = CSng(vsgrid.ColWidth(0)): intW = CSng(vsgrid.ColWidth(0) + vsgrid.ColWidth(1))
    If X > intX And X < intW And Y > 255 And Y < CSng(vsgrid.Rows * 255) And vsgrid.MouseRow > -1 Then
         vsgrid.MousePointer = flexCustom
         Set vsgrid.MouseIcon = Me.img16.ListImages(1).Picture
    Else
         vsgrid.MousePointer = flexDefault
         Set vsgrid.MouseIcon = Nothing
    End If
End Sub
'################################################################################################################
'## ���ܣ�  �������ļ���XML�ļ��е����������ļ��б���
'##
'##
'## ���أ�  ����ɹ�������Ture�����򷵻�False��
'################################################################################################################
Public Function ImportList() As Boolean
    '��XML�ļ�����
    Dim strXML As String, strArrXml As Variant, strTempName As String
    Dim i As Integer, j As Integer, k As Integer, l As Integer
    Dim oFileRootList As IXMLDOMNodeList      '�ļ��ڵ�
    Dim oFileRoot As IXMLDOMElement, strTemp As String
    Dim oRoot  As IXMLDOMElement        '���ڵ�
    Dim rsTemp As ADODB.Recordset
    Dim strUnitName As String
    Static intRow As Long
    On Error Resume Next
    dlgThis.MaxFileSize = 32767
    dlgThis.Filter = "*.XML|*.xml"
    dlgThis.DialogTitle = "��(֧�ֶ�ѡ)"
    dlgThis.CancelError = True
    dlgThis.flags = &H10& Or &H200& Or &H80000
    dlgThis.ShowOpen

'    dlgThis.Action = 1
    
    If Err.Number = 32755 Then Err.Clear: ImportList = False: Exit Function
    If dlgThis.FileTitle = "" Then
        strTempName = Split(dlgThis.Filename, "\")(UBound(Split(dlgThis.Filename, "\")))
        strArrXml = Split(Trim(strTempName), Chr(0)) '�����ļ������ַ���
        strArrXml(0) = Replace(dlgThis.Filename, strTempName, strArrXml(0) & "\")
    Else
       strTempName = Split(dlgThis.Filename, "\")(UBound(Split(dlgThis.Filename, "\")))
       strTempName = Replace(dlgThis.Filename, strTempName, "") & "," & Split(dlgThis.Filename, "\")(UBound(Split(dlgThis.Filename, "\")))
       strArrXml = Split(strTempName, ",")
    End If
    With vsgrid
        If vsgrid.Tag = "Export" Or vsgrid.Tag = "" Then
            .Clear
            .FixedRows = 1: .ExplorerBar = flexExSortShow
            .Cols = 6: .ColWidth(mImportCols.Choose) = 270: .Rows = 1: .ColAlignment(mImportCols.cName) = flexAlignLeftCenter
            .ColWidth(mImportCols.cImportType) = 0: .ColWidth(mImportCols.cName) = 1500: .ROWHEIGHT(mImportCols.Choose) = 50: .ColWidth(mImportCols.cTip) = 2500: .ColWidth(mImportCols.cUnit) = 2500: .ColWidth(mImportCols.cPath) = 6000
            .TextMatrix(0, mImportCols.Choose) = "ѡ��": .TextMatrix(0, mImportCols.cName) = "����": .TextMatrix(0, mImportCols.cTip) = "��ʾ": .TextMatrix(0, mImportCols.cUnit) = "������λ": .TextMatrix(0, mImportCols.cPath) = "�ļ�λ��"
            intRow = 0
        End If
    End With
    For k = 1 To UBound(strArrXml)
        strXML = strArrXml(0) & strArrXml(k)
        Set mdoc = New DOMDocument
        mdoc.Load strXML
        '�����·�����ļ��ѱ��������ټ���
        For l = 1 To vsgrid.Rows - 1
            If strXML = Trim(vsgrid.TextMatrix(l, mImportCols.cPath)) Then
                 MsgBox strArrXml(k) & ",�Ѿ����򿪣������ظ��� ��", vbInformation, gstrSysName
                 GoTo a
            End If
        Next l
        '����������κ�Ԫ�أ����˳�
        If mdoc.documentElement Is Nothing Then
           MsgBox "��ѡ���XML�ļ����Ǹ������������ȷXML��ʽ���ļ�!", vbInformation, gstrSysName: Exit Function
        End If
        '��ȡ�ļ��ṹ
        Set oRoot = mdoc.selectSingleNode("Document")       'oRoot��Ϊ���ڵ�
        Set oFileRootList = oRoot.selectNodes("File")
        If oRoot Is Nothing Then
            MsgBox "��ѡ���XML�ļ����Ǹ������������ȷXML��ʽ���ļ�!", vbInformation, gstrSysName: Exit Function
        ElseIf Not oRoot.selectSingleNode("EPRFileInfo") Is Nothing Then
             strTemp = strTemp & oRoot.selectSingleNode("EPRFileInfo").selectSingleNode("ID").Text & "_" & oRoot.selectSingleNode("EPRFileInfo").selectSingleNode("����").Text & "_" & oRoot.selectSingleNode("EPRFileInfo").selectSingleNode("����").Text & ","
        ElseIf oFileRootList.Item(0) Is Nothing Then
            MsgBox "��ѡ���XML�ļ����ݿ���Ϊ��!", vbInformation, gstrSysName: Exit Function
        Else
            For Each oFileRoot In oFileRootList
                strTemp = strTemp & GetNodeValue(oFileRoot, "ID", 0) & "_" & GetNodeValue(oFileRoot, "����", 0) & "_" & GetNodeValue(oFileRoot, "����", 0) & "_" & oRoot.getAttribute("UnitName") & ","
            Next
        End If
        strTemp = Mid(strTemp, 1, Len(strTemp) - 1)
        gstrSQL = "  select distinct ID,����,���,���� from (" & _
                  "  Select l.Id, decode(l.����,1,'1-���ﲡ��',2,'2-סԺ����',4,'4-������',5,'5-����֤������',6,'6-֪���ļ�') as ����, l.���, l.����" & _
                  "  From �����ļ��б� l where l.����<>2 and l.���� in (1,2,4,5,6)" & _
                  "  Union All Select 0, decode(l.����,1,'1-���ﲡ��',2,'2-סԺ����',4,'4-������',5,'5-����֤������',6,'6-֪���ļ�') as ����, null, null" & _
                  "  From �����ļ��б� l where l.����<>2 and l.���� in (1,2,4,5,6)) order by ����,ID"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        With vsgrid
            For i = 1 To UBound(Split(strTemp, ",")) + 1
                   Me.Tag = Val(Me.Tag) + 1
                   intRow = Me.Tag
                  .AddItem ""
                  .TextMatrix(intRow, mImportCols.cName) = Split(Split(strTemp, ",")(i - 1), "_")(1)
                  .TextMatrix(intRow, mImportCols.cImportType) = Split(Split(strTemp, ",")(i - 1), "_")(1) & "_2"
                  .TextMatrix(intRow, mImportCols.cTip) = "������,���뽫�������ļ���"
                  .TextMatrix(intRow, mImportCols.cPath) = strXML
                  .Cell(flexcpData, intRow, mImportCols.cName) = Split(Split(strTemp, ",")(i - 1), "_")(0)
                  .TextMatrix(intRow, mImportCols.cUnit) = Split(Split(strTemp, ",")(i - 1), "_")(3)
                  If Not rsTemp Is Nothing Then
                      Do While Not rsTemp.EOF
                          If Trim(NVL(rsTemp!����, "")) = Trim(Split(Split(strTemp, ",")(i - 1), "_")(1)) And Val(NVL(rsTemp!����, "")) = Split(Split(strTemp, ",")(i - 1), "_")(2) Then
                              .Cell(flexcpForeColor, intRow, mImportCols.cName, intRow, mImportCols.cTip) = vbMagenta
                              .Cell(flexcpData, intRow, mImportCols.cTip) = "����"
                              .TextMatrix(intRow, mImportCols.cImportType) = Split(Split(strTemp, ",")(i - 1), "_")(1) & "_1"
                              .TextMatrix(intRow, mImportCols.cTip) = "�Ѵ���,���뽫����ԭ���ļ���"
                              j = 1
                          End If
                          rsTemp.MoveNext
                      Loop
                  End If
                  '�ж��ļ��б����Ƿ�����ͬ�����ļ�����ͬ��ѡ��
                  If Not IsHave(intRow) Then
                    .Cell(flexcpPicture, intRow, mImportCols.Choose) = img16.ListImages("Check").Picture
                    .Cell(flexcpData, i, mImportCols.Choose) = 1
                  End If
                  vsgrid.Cell(flexcpData, intRow, mImportCols.Choose) = 1
                  rsTemp.MoveFirst
            Next i
        End With
        strTemp = ""
        Me.Tag = vsgrid.Rows - 1
        
a:
    Next k
    If j = 1 Then Me.ShowControl(Me.cbsThis, 110, True).Visible = True
    If vsgrid.Rows > 1 Then vsgrid.Row = 1
    dlgThis.Filename = ""
    ImportList = True
    Exit Function
errHand:
    ImportList = False
     If ErrCenter() = 1 Then Resume
       Call SaveErrLog
End Function
'################################################################################################################
'## ���ܣ�  ��ʼ�������ļ�������XML�ĵ���
'##
'## ������  eEdtMode    :��ǰ�༭ģʽ���������޸ģ�
'##         eEdtType    :��ǰ�༭��ʽ���ļ����塢ʾ���༭���������༭����������ˣ�
'##         lngArrFile  :��ǰѡ�����ID�����Ƽ���
'##       lngExportType :��������(1,��ʾ���������ļ���2��ʾ��������ļ�)
'## ���أ�  ����ɹ�������Ture�����򷵻�False��
'################################################################################################################
Public Function StartExportToXMLFile(ByVal eEdtMode As EditModeEnum, ByVal eEdtType As EditTypeEnum, ByVal lngArrFile, ByVal lngExportType As Long) As Boolean
    Dim i As Long, lngFileID As Long, strFileName As String, strFileType As String
    Dim cDoc As New DOMDocument              'xml�ĵ�
    Dim Result As VbMsgBoxResult
    Dim pi As IXMLDOMProcessingInstruction  '�汾��Ϣ
    Dim oRootNew As IXMLDOMElement
    Dim oRoot As IXMLDOMElement         '���ڵ�
    '------------------------------------------------
    
    If gobjFSO.FileExists(dlgThis.Filename) Then
        DoEvents
        If MsgBox(dlgThis.Filename & "�ļ��Ѿ����ڣ��Ƿ񸲸ǣ�", vbOKCancel + vbQuestion, gstrSysName) = vbCancel Then Exit Function
    End If
    cDoc.appendChild cDoc.createComment(gstrSysName & "����Ա:" & gstrUserName & "������:" & gstrDeptName & "��ʱ��:" & Format(Now(), "YYYY��MM��DD��"))
    Me.progBar.Visible = True
    For i = 0 To UBound(lngArrFile)
        DoEvents
        Me.Refresh
        strFileName = Split(lngArrFile(i), "_")(1)
        lngFileID = Split(lngArrFile(i), "_")(0)
        strFileType = Split(Split(lngArrFile(i), "_")(2), "-")(0)
        '��ͨסԺ����
        ZLCommFun.ShowFlash "���ڵ����ļ������Ժ�..."
        Screen.MousePointer = vbHourglass
        If lngExportType = 1 Then
             Call ExportToXml(cprEM_�޸�, cprET_�����ļ�����, lngFileID, cDoc, oRoot)
        ElseIf lngExportType = 2 Then
             Set cDoc = New DOMDocument
             cDoc.appendChild cDoc.createComment(gstrSysName & "����Ա:" & gstrUserName & "������:" & gstrDeptName & "��ʱ��:" & Format(Now(), "YYYY��MM��DD��"))
             Set oRootNew = Nothing
             Call ExportToXml(cprEM_�޸�, cprET_�����ļ�����, lngFileID, cDoc, oRootNew)
             Set pi = cDoc.createProcessingInstruction("xml", "version='1.0' encoding='gb2312'")
             Call cDoc.insertBefore(pi, cDoc.childNodes(0))
             cDoc.Save mstrPath & "/" & "����_" & strFileName & ".XML"
        End If
        Me.progBar.Value = IIf(Me.progBar.Value + Me.progBar.Max / (UBound(lngArrFile) + 1) > progBar.Max, progBar.Max, Me.progBar.Value + Me.progBar.Max / (UBound(lngArrFile) + 1))
    Next i
    If lngExportType <> 2 Then
        Set pi = cDoc.createProcessingInstruction("xml", "version='1.0' encoding='gb2312'")
        Call cDoc.insertBefore(pi, cDoc.childNodes(0))
        cDoc.Save mstrPath & "/" & zl9ComLib.GetUnitName & "_�����ļ��б�.XML"
        Set cDoc = Nothing
    End If
    Screen.MousePointer = vbDefault
    Me.progBar.Value = Me.progBar.Max
    Me.progBar.Visible = False
    Me.progBar.Value = 0
End Function

'��ʼ��Vsgrid
Private Function InitVsGrid(ByVal strMsg As String)
    vsgrid.Tag = ""
    With vsgrid
        .Clear: .Cols = 1: .Rows = 1: .FixedRows = 1
        .ROWHEIGHT(0) = vsgrid.Height
        .TextMatrix(0, 0) = strMsg
        .Cell(flexcpFontSize, 0, 0) = 20
        .Cell(flexcpAlignment, 0, 0) = flexAlignCenterCenter
        .Cell(flexcpBackColor, 0, 0) = vbWhite
    End With
    Me.Tag = ""
    vsgrid.ROWHEIGHT(0) = Me.ScaleHeight
End Function
'------------------------------------------------
   '���ܣ� ����CommandBars�˵�������������ʾ״̬
   '������ CommandBars�ؼ���toolId ,bolOn����
   '���أ� �ð�ť����
'------------------------------------------------
Public Function ShowControl(cbrObj As CommandBars, toolId As Long, blnOn As Boolean) As CommandBarControl
    Dim Control As CommandBarControl
    Dim ControlMenu As CommandBarControl
    '   ������
    Set Control = cbrObj.FindControl(, toolId, , True)
    If Not Control Is Nothing Then
        Control.Enabled = blnOn
    End If
  Set ShowControl = Control
End Function
'����ID�ַ������ù�������ʾ״̬
Public Function ShowControlEnabled(ByVal strControlID As String, ByVal blnOn As Boolean)
    Dim strArrId As Variant, i As Integer
    strArrId = Split(strControlID, ",")
    For i = 0 To UBound(strArrId)
        Call ShowControl(Me.cbsThis, Val(strArrId(i)), blnOn)
    Next i
End Function
'��ȡ���м��ص�XML�ļ�·��(�������ظ���)
Public Function getXmlPath() As String
    Dim i As Integer, j As Integer, strResult As String, strArr As Variant
    For i = 1 To vsgrid.Rows - 1
        If Not vsgrid.Cell(flexcpPicture, i, 0) Is Nothing Then
            strResult = strResult & vsgrid.TextMatrix(i, 5) & ","
        End If
    Next i
    strArr = Split(Mid(strResult, 1, Len(strResult) - 1), ",")
    strResult = ""
    For i = 0 To UBound(strArr)
        If InStr(strResult, strArr(i)) = 0 Then
        strResult = strResult & "," & strArr(i)
        End If
    Next i
    strResult = Mid(strResult, 2)
    getXmlPath = strResult
End Function
'################################################################################################################
'## ���ܣ�  �����ݿ�����д�뵽XML
'##
'## ������  eEdtMode    :��ǰ�༭ģʽ���������޸ģ�
'##         eEdtType    :��ǰ�༭��ʽ���ļ����塢ʾ���༭���������༭����������ˣ�
'##         lngFileID   :�ļ�ID�����ݱ༭��ʽ�Ĳ�ͬ�����Ա�ʾ�ļ�����ID������ID���߲��˲���ID��
'##         oDoc        :XML����
'##         oRoot       :XML���ڵ�
'################################################################################################################
Public Function ExportToXml(ByVal eEdtMode As EditModeEnum, ByVal eEdtType As EditTypeEnum, _
ByVal lngFileID As Long, ByRef oDoc As DOMDocument, ByRef oRoot As IXMLDOMElement) As Boolean
    Dim rs As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim oFileRoot As IXMLDOMElement      '�ļ��ڵ�
    Dim oNode As IXMLDOMNode            '���ڵ�
    Dim oSubNode1 As IXMLDOMNode        '�ӽڵ�
    Dim oSubNode2 As IXMLDOMNode        '�ӽڵ�
    Dim oSubNode3 As IXMLDOMNode        '�ӽڵ�
    Dim oSubNode4 As IXMLDOMNode        '�ӽڵ�
    Dim EPRFileInfoNode As IXMLDOMNode  '������Ϣ�ڵ�
    Dim CompendsoNode As IXMLDOMNode    '��ٽڵ�
    Dim ElementsNode As IXMLDOMNode     'Ҫ�ؽڵ�
    Dim PicturesNode As IXMLDOMNode     'ͼƬ�ڵ�
    Dim TablesNode As IXMLDOMNode       '���ڵ�
    Dim TableCells As IXMLDOMNode       '������ı����Ͻڵ�
    Dim TableElements As IXMLDOMNode    '�����Ҫ�ؼ��Ͻڵ�
    Dim TablePictures As IXMLDOMNode    '�����ͼƬ���Ͻڵ�
    Dim CellNode As IXMLDOMNode         '��Ԫ��ڵ�
    Dim ContentNode As IXMLDOMNode      '���ݽڵ�
    Dim oStream As New ADODB.Stream     '������
    Dim strPath As String               '��ʱ�ļ�Ŀ¼
    Dim strTemp As String               '��ʱ�ļ�
    Dim strPic As String                '��ʱͼƬ�ļ�
    Dim strHeadRtfFile As String        '��ʱҳü�ļ�
    Dim strFootRtfFile As String        '��ʱҳ���ļ�
    Dim strContextFile As String        '��ʱ�����ļ�
    Dim TempPic As New StdPicture, strTempPic As String
    Dim strObjArr As Variant
    On Error GoTo errHand
    strPath = IIf(Environ$("tmp") <> vbNullString, Environ$("tmp"), Environ$("temp"))
    '�ж��Ƿ�д�뵽ͬһ��XML��
    If oRoot Is Nothing Then
        Set oRoot = oDoc.createElement("Document")
        Set oDoc.documentElement = oRoot    '����Ϊ���ڵ�
        Call oRoot.setAttribute("UnitName", zl9ComLib.GetUnitName)
    End If
    '���ò����ļ��ڵ�
    Set oFileRoot = CreateNode(1, oRoot, "File", NODE_ELEMENT, "")
    Call oFileRoot.setAttribute("EditType", eEdtType)

    '�����ݿ���ȡ�����ļ�������Ϣ
    gstrSQL = "Select a.ID, a.����, a.���, a.����, a.˵��, a.ҳ��, a.����, a.ͨ��, b.���� As ҳ������, b.����, b.��ʽ, b.ҳü, b.ҳ�� " & _
                " From �����ļ��б� a, ����ҳ���ʽ b " & _
                " Where a.ҳ�� = b.��� And a.���� = b.���� And a.Id = [1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngFileID)
    Call oFileRoot.setAttribute("Name", NVL(rs("����")))
    'EPRFileInfo
    Set EPRFileInfoNode = CreateNode(1, oFileRoot, "EPRFileInfo", NODE_ELEMENT, "")
    CreateNode 2, EPRFileInfoNode, "ID", , lngFileID      '�ӽڵ�
    CreateNode 2, EPRFileInfoNode, "����", , NVL(rs("����"), 1)  '1-���ﲡ��;2-סԺ����;3-�����¼;4-������;5-����֤������;6-֪���ļ�;7-���Ʊ���;8-��������
    CreateNode 2, EPRFileInfoNode, "���", , NVL(rs("���"), 0)
    CreateNode 2, EPRFileInfoNode, "����", , NVL(rs("����"))
    CreateNode 2, EPRFileInfoNode, "˵��", , NVL(rs("˵��"))
    CreateNode 2, EPRFileInfoNode, "ҳ��", , NVL(rs("ҳ��"))
    CreateNode 2, EPRFileInfoNode, "����", , NVL(rs("����"), 0)
    CreateNode 2, EPRFileInfoNode, "ͨ��", , NVL(rs("ͨ��"), 0)
    CreateNode 2, EPRFileInfoNode, "����", , NVL(rs("����"), 0)
    CreateNode 2, EPRFileInfoNode, "ҳ������", , NVL(rs("ҳ������"))
    CreateNode 2, EPRFileInfoNode, "��ʽ", , NVL(rs("��ʽ"))
    CreateNode 2, EPRFileInfoNode, "ҳü", , NVL(rs("ҳü"))
    CreateNode 2, EPRFileInfoNode, "ҳ��", , NVL(rs("ҳ��"))
    '��ȡ��������RTF
    strContextFile = zlBlobRead(1, lngFileID)
    If strContextFile <> "" Then
       strTemp = zlFileUnzip(strContextFile)
       Me.RTbContext.LoadFile strTemp
       gobjFSO.DeleteFile strTemp
       gobjFSO.DeleteFile strContextFile, True
    End If
    '��ȡҳü�ļ���.RTF��
    strHeadRtfFile = zlBlobRead(12, NVL(rs("����"), 1) & "-" & NVL(rs("ҳ��")), App.Path & "\Head.rtf")
    If gobjFSO.FileExists(strHeadRtfFile) Then
        Me.RTbHeadText.LoadFile strHeadRtfFile             '��ȡ�ļ�
        gobjFSO.DeleteFile strHeadRtfFile, True            'ɾ����ʱ�ļ�
    End If
    CreateNode 2, EPRFileInfoNode, "ҳü�ļ�", , Replace(Me.RTbHeadText.TextRTF, "]]>", "]] >")
    '��ȡҳ���ļ���.RTF��
    strFootRtfFile = zlBlobRead(13, NVL(rs("����"), 1) & "-" & NVL(rs("ҳ��")), App.Path & "\Foot.rtf")
    If gobjFSO.FileExists(strFootRtfFile) Then
        Me.RTbHeadText.LoadFile strFootRtfFile              '��ȡ�ļ�
        gobjFSO.DeleteFile strFootRtfFile, True             'ɾ����ʱ�ļ�
    End If
    CreateNode 2, EPRFileInfoNode, "ҳ���ļ�", , Replace(Me.RTbHeadText.TextRTF, "]]>", "]] >")
    '��ȡҳüͼƬ����
    strTempPic = zlBlobRead(7, NVL(rs("����"), 1) & "-" & NVL(rs("ҳ��")))
    If gobjFSO.FileExists(strTempPic) Then
        Set TempPic = LoadPicture(strTempPic)
        gobjFSO.DeleteFile strTempPic, True      'ɾ����ʱ�ļ�
        If Not TempPic Is Nothing Then
            oStream.Type = adTypeBinary
            oStream.Open
            strPic = strPath & "\XMLPIC" & App.hInstance & ".jpg"
            SavePicture TempPic, strPic
            oStream.LoadFromFile strPic
            Set oSubNode1 = oDoc.createElement("OrigPic")
            oSubNode1.datatype = "bin.base64"
            oSubNode1.nodeTypedValue = oStream.Read
            EPRFileInfoNode.appendChild oSubNode1
            oStream.Close
            If gobjFSO.FileExists(strPic) Then gobjFSO.DeleteFile strPic, True
        End If
    End If
    '��ȡԪ�ؼ���
    gstrSQL = "Select Level, ID, �ļ�id, ��id, �������, ��������, ������, ��������, ��������, �����д�, �����ı�, �Ƿ���, Ԥ�����id,�������ID, �������, ʹ��ʱ��," & vbNewLine & _
                "       ����Ҫ��id, �滻��, Ҫ������, Ҫ������, Ҫ�س���, Ҫ��С��, Ҫ�ص�λ, Ҫ�ر�ʾ, ������̬, Ҫ��ֵ��" & vbNewLine & _
                "From (Select ID, �ļ�id, ��id, �������, ��������, ������, ��������, ��������, �����д�, �����ı�, �Ƿ���,Ԥ�����id,ID �������ID,�������,ʹ��ʱ��," & vbNewLine & _
                "              ����Ҫ��id, �滻��, Ҫ������, Ҫ������, Ҫ�س���, Ҫ��С��, Ҫ�ص�λ, Ҫ�ر�ʾ, ������̬, Ҫ��ֵ��" & vbNewLine & _
                "       From �����ļ��ṹ" & vbNewLine & _
                "       Where �ļ�id = [1] And ������� > 0)" & vbNewLine & _
                "Start With ��id Is Null" & vbNewLine & _
                "Connect By Prior ID = ��id" & vbNewLine & _
                "Order By �������, �����д�"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngFileID)
     Do While Not rs.EOF
        Select Case NVL(rs("��������"), 2)
            Case 1  'Compends ��ٽڵ�
                 If CompendsoNode Is Nothing Then Set CompendsoNode = CreateNode(1, oFileRoot, "Compends", NODE_ELEMENT, "")
                 Set oSubNode1 = CreateNode(2, CompendsoNode, "Compend", NODE_ELEMENT, "")
                    CreateNode 3, oSubNode1, "Key", , NVL(rs!������, "")
                    CreateNode 3, oSubNode1, "ID", , rs!ID
                    CreateNode 3, oSubNode1, "�ļ�ID", , NVL(rs!�ļ�ID, 0)
                    CreateNode 3, oSubNode1, "��ID", , 0
                    CreateNode 3, oSubNode1, "�������", , NVL(rs!�������, 0)
                    CreateNode 3, oSubNode1, "��������", , IIf(NVL(rs!��������, 0) = 0, False, True)
                    CreateNode 3, oSubNode1, "����", , NVL(rs!�����ı�)
                    CreateNode 3, oSubNode1, "˵��", , NVL(rs!��������)
                    CreateNode 3, oSubNode1, "Ԥ�����ID", , NVL(rs!Ԥ�����ID, 0)
                    CreateNode 3, oSubNode1, "�������ID", , NVL(rs!�������ID, 0)
                    CreateNode 3, oSubNode1, "�������", , IIf(NVL(rs!�������, 0) = 0, False, True)
                    CreateNode 3, oSubNode1, "ʹ��ʱ��", , NVL(rs!ʹ��ʱ��)
                    CreateNode 3, oSubNode1, "Level", , NVL(rs!Level, 0)
                    CreateNode 3, oSubNode1, "�ڲ����", , NVL(rs!�������, 0)
            Case 3  '���
                If TablesNode Is Nothing Then Set TablesNode = CreateNode(1, oFileRoot, "Tables", NODE_ELEMENT, "")
                Set oSubNode1 = CreateNode(2, TablesNode, "Table", NODE_ELEMENT, "")
                    CreateNode 3, oSubNode1, "Key", , NVL(rs!������, "")
                    CreateNode 3, oSubNode1, "ID", , NVL(rs!ID, 0)
                    CreateNode 3, oSubNode1, "�ļ�ID", , NVL(rs!�ļ�ID, 0)
                    CreateNode 3, oSubNode1, "��ID", , NVL(rs!��ID, 0)
                    CreateNode 3, oSubNode1, "�������", , NVL(rs!�������, 0)
                    CreateNode 3, oSubNode1, "��������", , IIf(NVL(rs!��������, 0) = 0, False, True)
                    CreateNode 3, oSubNode1, "�Ƿ���", , IIf(NVL(rs!�Ƿ���, 0) = 0, False, True)
                    CreateNode 3, oSubNode1, "Ԥ�����ID", , NVL(rs!Ԥ�����ID, 0)
                    CreateNode 3, oSubNode1, "��������", , NVL(rs!��������)
                    gstrSQL = "Select Level, ID, �ļ�id, ��id, �������, ��������, ������, ��������, ��������, �����д�, �����ı�, �Ƿ���, Ԥ�����id," & _
                              "����Ҫ��ID , �滻��, Ҫ������, Ҫ������, Ҫ�س���, Ҫ��С��, Ҫ�ص�λ, Ҫ�ر�ʾ, ������̬, Ҫ��ֵ�� " & _
                              "From (Select ID, �ļ�id, ��id, �������, ��������, ������, ��������, ��������, �����д�, �����ı�, �Ƿ���, Ԥ�����id, ����Ҫ��id, " & _
                              "�滻�� , Ҫ������, Ҫ������, Ҫ�س���, Ҫ��С��, Ҫ�ص�λ, Ҫ�ر�ʾ, ������̬, Ҫ��ֵ�� From �����ļ��ṹ " & _
                              "Where �ļ�id = " & NVL(rs!�ļ�ID, 0) & " ) Start With ��id + 0 = " & NVL(rs!ID, 0) & " Connect By Prior ID = ��id + 0 Order By ������, �������, �����д�"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
                    Set TableCells = CreateNode(2, oSubNode1, "Cells", NODE_ELEMENT, "")
                    Set TableElements = CreateNode(2, oSubNode1, "Elements", NODE_ELEMENT, "")
                    Set TablePictures = CreateNode(2, oSubNode1, "Pictures", NODE_ELEMENT, "")
                    Do While Not rsTemp.EOF
                        Select Case NVL(rsTemp!��������, 0)
                            Case 2  '�ı�
                                Set CellNode = CreateNode(3, TableCells, "Cell", NODE_ELEMENT, "")
                                    CreateNode 4, CellNode, "Key", , NVL(rsTemp!������, "")
                                    CreateNode 4, CellNode, "ID", , NVL(rsTemp!ID, 0)
                                    CreateNode 4, CellNode, "�ļ�ID", , NVL(rsTemp!�ļ�ID, 0)
                                    CreateNode 4, CellNode, "��ID", , NVL(rsTemp!��ID, 0)
                                    CreateNode 4, CellNode, "�������", , NVL(rsTemp!�������, 0)
                                    CreateNode 4, CellNode, "�����ı�", , NVL(rsTemp!�����ı�, "")
                                    CreateNode 4, CellNode, "��������", , IIf(NVL(rsTemp!��������, 0) = 0, False, True)
                                    CreateNode 4, CellNode, "��������", , NVL(rsTemp!��������)
                            Case 4  '����Ҫ��
                                 '�����ӵ�Ԫ Cell �� Cells���Ͻڵ���
                                 Set CellNode = CreateNode(3, TableCells, "Cell", NODE_ELEMENT, "")
                                    CreateNode 4, CellNode, "Key", , NVL(rsTemp!������, "")
                                    CreateNode 4, CellNode, "ID", , NVL(rsTemp!ID, 0)
                                    CreateNode 4, CellNode, "�ļ�ID", , NVL(rsTemp!�ļ�ID, 0)
                                    CreateNode 4, CellNode, "��ID", , NVL(rsTemp!��ID, 0)
                                    CreateNode 4, CellNode, "�������", , NVL(rsTemp!�������, 0)
                                    CreateNode 4, CellNode, "�����ı�", , NVL(rsTemp!�����ı�, "")
                                    CreateNode 4, CellNode, "��������", , IIf(NVL(rsTemp!��������, 0) = 0, False, True)
                                    CreateNode 4, CellNode, "��������", , NVL(rsTemp!��������)
                                 '����������Ҫ�� Element �� Elements���Ͻڵ���
                                Set oSubNode3 = CreateNode(3, TableElements, "Element", NODE_ELEMENT, "")
                                    CreateNode 4, oSubNode3, "Key", , NVL(rsTemp!������, "")
                                    CreateNode 4, oSubNode3, "ID", , rsTemp!ID
                                    CreateNode 4, oSubNode3, "�ļ�ID", , NVL(rsTemp!�ļ�ID, 0)
                                    CreateNode 4, oSubNode3, "��ID", , NVL(rsTemp!��ID, 0)
                                    CreateNode 4, oSubNode3, "�������", , NVL(rsTemp!�������, 0)
                                    CreateNode 4, oSubNode3, "��������", , IIf(NVL(rsTemp!��������, 0) = 0, False, True)
                                    CreateNode 4, oSubNode3, "�����ı�", , NVL(rsTemp!�����ı�)
                                    CreateNode 4, oSubNode3, "�Ƿ���", , IIf(NVL(rsTemp!�Ƿ���, 0) = 0, False, True)
                                    CreateNode 4, oSubNode3, "����Ҫ��ID", , NVL(rsTemp!����Ҫ��ID, 0)
                                    CreateNode 4, oSubNode3, "�滻��", , NVL(rsTemp!�滻��, 0)
                                    CreateNode 4, oSubNode3, "Ҫ������", , NVL(rsTemp!Ҫ������)
                                    CreateNode 4, oSubNode3, "Ҫ������", , NVL(rsTemp!Ҫ������, 0)
                                    CreateNode 4, oSubNode3, "Ҫ�س���", , NVL(rsTemp!Ҫ�س���, 0)
                                    CreateNode 4, oSubNode3, "Ҫ��С��", , NVL(rsTemp!Ҫ��С��, 0)
                                    CreateNode 4, oSubNode3, "Ҫ�ص�λ", , NVL(rsTemp!Ҫ�ص�λ)
                                    CreateNode 4, oSubNode3, "Ҫ�ر�ʾ", , NVL(rsTemp!Ҫ�ر�ʾ, 0)
                                    CreateNode 4, oSubNode3, "������̬", , NVL(rsTemp!������̬, 0)
                                    CreateNode 4, oSubNode3, "Ҫ��ֵ��", , NVL(rsTemp!Ҫ��ֵ��)
                                    CreateNode 4, oSubNode3, "��������", , NVL(rsTemp!��������)
                            Case 5  'ͼƬ
                                 Set oSubNode3 = CreateNode(3, TablePictures, "Picture", NODE_ELEMENT, "")
                                    CreateNode 4, oSubNode3, "Key", , NVL(rsTemp!������, "")
                                    CreateNode 4, oSubNode3, "ID", , NVL(rsTemp!ID, 0)
                                    CreateNode 4, oSubNode3, "�ļ�ID", , NVL(rsTemp!�ļ�ID, 0)
                                    CreateNode 4, oSubNode3, "��ID", , NVL(rsTemp!��ID, 0)
                                    CreateNode 4, oSubNode3, "�������", , NVL(rsTemp!�������, 0)
                                    CreateNode 4, oSubNode3, "��������", , IIf(NVL(rsTemp!��������, 0) = 0, False, True)
                                    CreateNode 4, oSubNode3, "�����ı�", , NVL(rsTemp!�����ı�, "")
                                    CreateNode 4, oSubNode3, "�Ƿ���", , IIf(NVL(rsTemp!�Ƿ���, 0) = 0, False, True)
                                    CreateNode 4, oSubNode3, "��������", , NVL(rsTemp!��������, "")
                                    '�洢ͼƬ����
                                    strTempPic = zlBlobRead(2, rsTemp!ID)
                                    Set TempPic = LoadPicture(strTempPic)
                                    gobjFSO.DeleteFile strTempPic, True      'ɾ����ʱ�ļ�
                                    oStream.Type = adTypeBinary
                                    oStream.Open
                                    strPic = strPath & "\XMLPIC" & App.hInstance & ".jpg"
                                    SavePicture TempPic, strPic
                                    oStream.LoadFromFile strPic
                                    Set oSubNode4 = oDoc.createElement("OrigPic")
                                    oSubNode4.datatype = "bin.base64"
                                    oSubNode4.nodeTypedValue = oStream.Read
                                    oSubNode3.appendChild oSubNode4
                                    oStream.Close
                                    'ɾ����ʱ�ļ�
                                    If gobjFSO.FileExists(strPic) Then gobjFSO.DeleteFile strPic, True
                        End Select
                        rsTemp.MoveNext
                    Loop
            Case 4  'Ҫ��
                 If ElementsNode Is Nothing Then Set ElementsNode = CreateNode(1, oFileRoot, "Elements", NODE_ELEMENT, "")
                 Set oSubNode1 = CreateNode(2, ElementsNode, "Element", NODE_ELEMENT, "")
                    CreateNode 3, oSubNode1, "Key", , NVL(rs!������, "")
                    CreateNode 3, oSubNode1, "ID", , rs!ID
                    CreateNode 3, oSubNode1, "�ļ�ID", , NVL(rs!�ļ�ID, 0)
                    CreateNode 3, oSubNode1, "��ID", , NVL(rs!��ID, 0)
                    CreateNode 3, oSubNode1, "�������", , NVL(rs!�������, 0)
                    CreateNode 3, oSubNode1, "��������", , IIf(NVL(rs!��������, 0) = 0, False, True)
                    CreateNode 3, oSubNode1, "�����ı�", , NVL(rs!�����ı�)
                    CreateNode 3, oSubNode1, "�Ƿ���", , IIf(NVL(rs!�Ƿ���, 0) = 0, False, True)
                    CreateNode 3, oSubNode1, "����Ҫ��ID", , NVL(rs!����Ҫ��ID, 0)
                    CreateNode 3, oSubNode1, "�滻��", , NVL(rs!�滻��, 0)
                    CreateNode 3, oSubNode1, "Ҫ������", , NVL(rs!Ҫ������)
                    CreateNode 3, oSubNode1, "Ҫ������", , NVL(rs!Ҫ������, 0)
                    CreateNode 3, oSubNode1, "Ҫ�س���", , NVL(rs!Ҫ�س���, 0)
                    CreateNode 3, oSubNode1, "Ҫ��С��", , NVL(rs!Ҫ��С��, 0)
                    CreateNode 3, oSubNode1, "Ҫ�ص�λ", , NVL(rs!Ҫ�ص�λ)
                    CreateNode 3, oSubNode1, "Ҫ�ر�ʾ", , NVL(rs!Ҫ�ر�ʾ, 0)
                    CreateNode 3, oSubNode1, "������̬", , NVL(rs!������̬, 0)
                    CreateNode 3, oSubNode1, "Ҫ��ֵ��", , NVL(rs!Ҫ��ֵ��)
                    CreateNode 3, oSubNode1, "��������", , NVL(rs!��������)
            Case 5  'ͼƬ
                 If PicturesNode Is Nothing Then Set PicturesNode = CreateNode(1, oFileRoot, "Pictures", NODE_ELEMENT, "")
                 Set oSubNode1 = CreateNode(2, PicturesNode, "Picture", NODE_ELEMENT, "")
                    CreateNode 3, oSubNode1, "Key", , NVL(rs!������, "")
                    CreateNode 3, oSubNode1, "ID", , NVL(rs!ID, 0)
                    CreateNode 3, oSubNode1, "�ļ�ID", , NVL(rs!�ļ�ID, 0)
                    CreateNode 3, oSubNode1, "��ID", , NVL(rs!��ID, 0)
                    CreateNode 3, oSubNode1, "�������", , NVL(rs!�������, 0)
                    CreateNode 3, oSubNode1, "��������", , IIf(NVL(rs!��������, 0) = 0, False, True)
                    CreateNode 3, oSubNode1, "�����ı�", , NVL(rs!�����ı�)
                    CreateNode 3, oSubNode1, "�Ƿ���", , IIf(NVL(rs!�Ƿ���, 0) = 0, False, True)
                    CreateNode 3, oSubNode1, "��������", , NVL(rs!��������, "")
                    '�洢ͼƬ����
                    strTempPic = zlBlobRead(2, rs!ID)
                    If strTempPic <> "" Then
                        Set TempPic = LoadPicture(strTempPic)
                        gobjFSO.DeleteFile strTempPic, True      'ɾ����ʱ�ļ�
                        oStream.Type = adTypeBinary
                        oStream.Open
                        strPic = strPath & "\XMLPIC" & App.hInstance & ".jpg"
                        SavePicture TempPic, strPic
                        oStream.LoadFromFile strPic
                        Set oSubNode2 = oDoc.createElement("OrigPic")
                        oSubNode2.datatype = "bin.base64"
                        oSubNode2.nodeTypedValue = oStream.Read
                        oSubNode1.appendChild oSubNode2
                        oStream.Close
                        'ɾ����ʱ�ļ�
                        If gobjFSO.FileExists(strPic) Then gobjFSO.DeleteFile strPic, True
                    End If
          End Select
        rs.MoveNext
    Loop
     'RTF�ı�
    Set oNode = CreateNode(1, oFileRoot, "Content", NODE_ELEMENT, "")
    Set oSubNode1 = CreateNode(2, oNode, "RTF", NODE_ELEMENT, "")
    CreateNode 3, oSubNode1, "RTFText", NODE_CDATA_SECTION, Replace(Me.RTbContext.TextRTF, "]]>", "]] >")
    Exit Function
errHand:
    ExportToXml = False
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    ExportToXml = True
End Function
'################################################################################################################
'## ���ܣ�  ��XML���ݵ��뵽���ݿ�
'##
'## ������  strFilePath    :XML�ļ���·��
'##         strFileName    :���벡���ļ������ƣ�����XML�����ƣ�
'################################################################################################################
Private Function ImportFromXml(ByVal strFilePath As String, ByVal strFileName As String) As Boolean
    '---------------------------------------------------
    Dim oDoc As New DOMDocument
    Dim oRoot As IXMLDOMElement         '���ڵ�
    Dim oFileRoot As IXMLDOMElement     '�ļ��ڵ�
    Dim oNodeList As IXMLDOMNodeList    '�ڵ㼯��
    Dim oNode As IXMLDOMNode            '�ӽڵ�
    Dim oSubNode1 As IXMLDOMNode        '�ӽڵ�
    Dim oSubNode2 As IXMLDOMNode        '�ӽڵ�
    Dim oSubNode3 As IXMLDOMNode        '�ӽڵ�
    Dim oSubNode4 As IXMLDOMNode        '�ӽڵ�
    Dim EPRFileInfoNode As IXMLDOMNode  '������Ϣ�ڵ�
    Dim Compends As IXMLDOMNodeList     '��ٽڵ�
    Dim Elements As IXMLDOMNodeList     'Ҫ�ؽڵ�
    Dim Pictures As IXMLDOMNodeList     'ͼƬ�ڵ�
    Dim Tables As IXMLDOMNodeList       '���ڵ�
    Dim Cells As IXMLDOMNodeList        '������ı����Ͻڵ�
    Dim TableElements As IXMLDOMNodeList    '�����Ҫ�ؼ��Ͻڵ�
    Dim TablePictures As IXMLDOMNodeList    '�����ͼƬ���Ͻڵ�
    Dim CellNode As IXMLDOMNode         '��Ԫ��ڵ�
    Dim ContentNode As IXMLDOMNode      '���ݽڵ�
    Dim oStream As New ADODB.Stream     '������
    Dim strPath As String               '��ʱ�ļ�Ŀ¼
    Dim strTemp As String               '��ʱ�ļ�
    Dim strPic As String                '��ʱͼƬ�ļ�
    Dim strHeadRtfFile As String        '��ʱҳü�ļ�
    Dim strFootRtfFile As String        '��ʱҳ���ļ�
    Dim strContextFile As String        '��ʱ������
    Dim strArrNames As Variant          '�ļ���������
    Dim ArraySQL() As String            'SQL����
    '-------------------------------------------------------------
    Dim strTempName As String
    
    Dim lngTempID As Long
    Dim lngCompendID As Long, lngID As Long, lng�д� As Long
    '-------------------------------------------------------------
    Dim GpInput As GdiplusStartupInput
    Dim m_GDIpToken         As Long         ' ���ڹر� GDI+
    Dim oDIB As New cDIB
    Dim DIBDither As New cDIBDither
    Dim DIBPal As New cDIBPal
    '-------------------------------------------------------------
    Dim TempPic As New StdPicture, strTempPic As String
    Dim rsTemp As New ADODB.Recordset, rs As New ADODB.Recordset
    Dim Result As VbMsgBoxResult
    '-------------------------------------------------------------
    On Error GoTo errHand
    'oDoc.Load strFilePath
    Set oDoc = mdoc
    Set oRoot = oDoc.selectSingleNode("Document")
    If oRoot Is Nothing Then GoTo errMsg
    strArrNames = Split(strFileName, "_")
    Set oFileRoot = oRoot.selectSingleNode("/Document/File[@Name='" & strArrNames(0) & "']")
    If oFileRoot Is Nothing Then
        If Not oRoot.selectSingleNode("EPRFileInfo") Is Nothing Then
            Set oFileRoot = oRoot
        End If
    End If
    ReDim ArraySQL(1 To 1) As String
    '���뷽ʽ�ж�
    Set EPRFileInfoNode = oFileRoot.selectSingleNode("EPRFileInfo")
    strTempName = EPRFileInfoNode.selectSingleNode("����").Text
    lngTempID = NVL(EPRFileInfoNode.selectSingleNode("ID").Text, 0)
    Dim strPageName As String
    If Val(strArrNames(1)) = 2 Then '���뷽ʽ��1.���� �� 2.������
         Dim strSearchName As String
         gstrSQL = "select ���� from �����ļ��б� where ���� like '" & strTempName & "%'"
         Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "")
         If rsTemp.RecordCount = 0 Then
             EPRFileInfoNode.selectSingleNode("����").Text = strTempName
             strPageName = strTempName
         ElseIf rsTemp.RecordCount = 1 Then
             If rsTemp!���� = EPRFileInfoNode.selectSingleNode("����").Text Then
                EPRFileInfoNode.selectSingleNode("����").Text = strTempName & "-1"
                strPageName = strTempName & "-1"
             End If
         Else
             gstrSQL = "select '" & strTempName & "-'|| max(to_number(replace(����,'" & strTempName & "-',''))+1) as ���� " & _
             " from �����ļ��б� where ���� like '" & strTempName & "-%' and instr(replace(����,'" & strTempName & "-',''),'-')<1"
             Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "")
             EPRFileInfoNode.selectSingleNode("����").Text = rsTemp!����
             strPageName = rsTemp!����
         End If
         Set rsTemp = zlDatabase.OpenSQLRecord("select LPad(nvl(max(���),0)+1, 3,0 ) as ��� from �����ļ��б� Where ���� = [1]", "����", EPRFileInfoNode.selectSingleNode("����").Text)
         EPRFileInfoNode.selectSingleNode("ҳ��").Text = rsTemp!���
         Set rsTemp = zlDatabase.OpenSQLRecord("Select nvl(nvl(max(ID),0)+1,'000') as ID From �����ļ��б�", "ID")
         lngTempID = Val(rsTemp!ID)   '����ID
         gstrSQL = "Zl_�����ļ��б�_Insert('" & lngTempID & "','" & EPRFileInfoNode.selectSingleNode("����").Text & "','" & EPRFileInfoNode.selectSingleNode("ҳ��").Text & "','" & EPRFileInfoNode.selectSingleNode("����").Text & "','" & EPRFileInfoNode.selectSingleNode("˵��").Text & "','" & EPRFileInfoNode.selectSingleNode("ҳ��").Text & "','" & strPageName & "','" & EPRFileInfoNode.selectSingleNode("����").Text & "')"
         ReDim Preserve ArraySQL(1 To UBound(ArraySQL) + 1) As String
         ArraySQL(UBound(ArraySQL)) = gstrSQL
    Else
         gstrSQL = "select a.ID,a.����,a.���,a.ҳ�� ,b.���� as ����ҳ�� from �����ļ��б� a , �����ļ��б� b " & _
              " Where a.���� = " & EPRFileInfoNode.selectSingleNode("����").Text & " and b.����=" & EPRFileInfoNode.selectSingleNode("����").Text & " And b.��� = a.ҳ�� And a.���� ='" & EPRFileInfoNode.selectSingleNode("����").Text & "'"
         Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "")
         lngTempID = Val(rsTemp!ID)
         EPRFileInfoNode.selectSingleNode("ID").Text = Val(rsTemp!ID)
         EPRFileInfoNode.selectSingleNode("ҳ��").Text = rsTemp!ҳ��
         strPageName = rsTemp!����ҳ��
    End If
    '��XML��ȡ�ļ�������Ϣ
    Set EPRFileInfoNode = oFileRoot.selectSingleNode("EPRFileInfo")
    ReDim Preserve ArraySQL(1 To UBound(ArraySQL) + 1) As String
    ArraySQL(UBound(ArraySQL)) = "Zl_����ҳ���ʽ_Update(" & EPRFileInfoNode.selectSingleNode("����").Text & ",'" & EPRFileInfoNode.selectSingleNode("ҳ��").Text & "','" & strPageName & "'," & _
    EPRFileInfoNode.selectSingleNode("����").Text & ",'" & EPRFileInfoNode.selectSingleNode("��ʽ").Text & "'," & _
    "'" & EPRFileInfoNode.selectSingleNode("ҳü").Text & "','" & EPRFileInfoNode.selectSingleNode("ҳ��").Text & "')"
    '��XML��ȡҳü�ļ�
    If Not EPRFileInfoNode.selectSingleNode("ҳü�ļ�") Is Nothing Then Me.RTbHeadText.TextRTF = EPRFileInfoNode.selectSingleNode("ҳü�ļ�").Text
    If Me.RTbHeadText.TextRTF <> "" Then
        Me.RTbHeadText.SaveFile App.Path & "\Head.rtf"
        Call zlBlobSql(12, EPRFileInfoNode.selectSingleNode("����").Text & "-" & EPRFileInfoNode.selectSingleNode("ҳ��").Text, App.Path & "\Head.rtf", ArraySQL)
        gobjFSO.DeleteFile App.Path & "\Head.rtf", True
    End If
    '��XML��ȡҳ���ļ�
    If Not EPRFileInfoNode.selectSingleNode("ҳ���ļ�") Is Nothing Then Me.RTbFootText.TextRTF = EPRFileInfoNode.selectSingleNode("ҳ���ļ�").Text
    If Me.RTbFootText.TextRTF <> "" Then
        Me.RTbFootText.SaveFile App.Path & "\Foot.rtf"
        Call zlBlobSql(13, EPRFileInfoNode.selectSingleNode("����").Text & "-" & EPRFileInfoNode.selectSingleNode("ҳ��").Text, App.Path & "\Foot.rtf", ArraySQL)
        gobjFSO.DeleteFile App.Path & "\Foot.rtf", True
    End If
    '��XML��ȡҳüͼƬ
    If Not EPRFileInfoNode.selectSingleNode("OrigPic") Is Nothing Then
        oStream.Type = adTypeBinary
        oStream.Open
        oStream.Write EPRFileInfoNode.selectSingleNode("OrigPic").nodeTypedValue
        strPic = App.Path & "\XML2JPG" & App.hInstance & ".JPG"
        oStream.SaveToFile strPic, adSaveCreateOverWrite
        oStream.Close
        Call zlBlobSql(7, EPRFileInfoNode.selectSingleNode("����").Text & "-" & EPRFileInfoNode.selectSingleNode("ҳ��").Text, strPic, ArraySQL)
        gobjFSO.DeleteFile strPic, True      'ɾ����ʱ�ļ�
    End If
    '��XML��ȡ�����Ϣ
    If Not oFileRoot.selectSingleNode("Compends") Is Nothing Then Set Compends = oFileRoot.selectSingleNode("Compends").selectNodes("Compend")
    If Not Compends Is Nothing Then
        For Each oNode In Compends
            lngCompendID = zlDatabase.GetNextId("�����ļ��ṹ")
            ReDim Preserve ArraySQL(1 To UBound(ArraySQL) + 1) As String
            ArraySQL(UBound(ArraySQL)) = "Zl_�����ļ��ṹ_Update(" & lngCompendID & "," & lngTempID & "," & IIf(oNode.selectSingleNode("��ID").Text = 0, "NULL", oNode.selectSingleNode("��ID").Text) & "," & _
                oNode.selectSingleNode("�������").Text & ",1," & oNode.selectSingleNode("Key").Text & "," & IIf(oNode.selectSingleNode("��������").Text, 1, 0) & ",'" & oNode.selectSingleNode("˵��").Text & "',NULL,'" & oNode.selectSingleNode("����").Text & "',NULL," & _
                IIf(oNode.selectSingleNode("Ԥ�����ID").Text = 0, "NULL", oNode.selectSingleNode("Ԥ�����ID").Text) & "," & IIf(oNode.selectSingleNode("�������").Text, 1, 0) & ",'" & oNode.selectSingleNode("ʹ��ʱ��").Text & "')"
            Set oNodeList = oFileRoot.selectNodes("//*[��ID=" & oNode.selectSingleNode("ID").Text & " ]")
            For Each oSubNode1 In oNodeList
                oSubNode1.selectSingleNode("��ID").Text = lngCompendID
            Next
        Next
       
    End If
    '��XML��ȡҪ����Ϣ
    If Not oFileRoot.selectSingleNode("Elements") Is Nothing Then Set Elements = oFileRoot.selectSingleNode("Elements").selectNodes("Element")
    If Not Elements Is Nothing Then
        For Each oNode In Elements
            ReDim Preserve ArraySQL(1 To UBound(ArraySQL) + 1) As String
            ArraySQL(UBound(ArraySQL)) = "Zl_�����ļ��ṹ_Update(" & zlDatabase.GetNextId("�����ļ��ṹ") & "," & lngTempID & "," & IIf(oNode.selectSingleNode("��ID").Text = 0, "NULL", oNode.selectSingleNode("��ID").Text) & "," & _
                oNode.selectSingleNode("�������").Text & ",4," & oNode.selectSingleNode("Key").Text & "," & IIf(oNode.selectSingleNode("��������").Text, 1, 0) & ",'" & oNode.selectSingleNode("��������").Text & "',NULL,'" & _
                Replace(oNode.selectSingleNode("�����ı�").Text, "'", "' || chr(39) || '") & "'," & IIf(oNode.selectSingleNode("�Ƿ���").Text, 1, 0) & ",NULL,NULL,NULL," & _
                IIf(CheckValid(oNode.selectSingleNode("����Ҫ��ID").Text, oNode.selectSingleNode("Ҫ������").Text), oNode.selectSingleNode("����Ҫ��ID").Text, "NULL") & "," & _
                oNode.selectSingleNode("�滻��").Text & ",'" & oNode.selectSingleNode("Ҫ������").Text & "'," & oNode.selectSingleNode("Ҫ������").Text & "," & oNode.selectSingleNode("Ҫ�س���").Text & "," & _
                oNode.selectSingleNode("Ҫ��С��").Text & ",'" & oNode.selectSingleNode("Ҫ�ص�λ").Text & "'," & oNode.selectSingleNode("Ҫ�ر�ʾ").Text & "," & oNode.selectSingleNode("������̬").Text & ",'" & oNode.selectSingleNode("Ҫ��ֵ��").Text & "')"
        Next
    End If
    '��XML��ȡ�����Ϣ
    If Not oFileRoot.selectSingleNode("Tables") Is Nothing Then Set Tables = oFileRoot.selectSingleNode("Tables").selectNodes("Table")
    If Not Tables Is Nothing Then
        For Each oNode In Tables
            lngID = zlDatabase.GetNextId("�����ļ��ṹ")
            '������ṹSQL���
            ReDim Preserve ArraySQL(1 To UBound(ArraySQL) + 1) As String
            ArraySQL(UBound(ArraySQL)) = "Zl_�����ļ��ṹ_Update(" & lngID & "," & lngTempID & "," & IIf(oNode.selectSingleNode("��ID").Text = 0, "NULL", oNode.selectSingleNode("��ID").Text) & "," & _
            oNode.selectSingleNode("�������").Text & ",3," & oNode.selectSingleNode("Key").Text & "," & IIf(oNode.selectSingleNode("��������").Text, 1, 0) & ",'" & oNode.selectSingleNode("��������").Text & "',NULL,'" & "" & "'," & IIf(oNode.selectSingleNode("�Ƿ���").Text, 1, 0) & _
            "," & IIf(oNode.selectSingleNode("Ԥ�����ID").Text = 0, "NULL", oNode.selectSingleNode("Ԥ�����ID").Text) & ")"
            '������������ĸ�ID
            Set oNodeList = oNode.selectNodes("//*[��ID=" & oNode.selectSingleNode("ID").Text & " ]")
            For Each oSubNode1 In oNodeList
                oSubNode1.selectSingleNode("��ID").Text = lngID
            Next
            '��ȡ����е�Ԫ��
            If Not oNode.selectSingleNode("Cells") Is Nothing Then Set Cells = oNode.selectSingleNode("Cells").selectNodes("Cell")
            If Not oNode.selectSingleNode("Elements") Is Nothing Then Set TableElements = oNode.selectSingleNode("Elements").selectNodes("Element")
            If Not oNode.selectSingleNode("Pictures") Is Nothing Then Set TablePictures = oNode.selectSingleNode("Pictures").selectNodes("Picture")
            '��Ԫ���ı���Ҫ��
            If Not Cells Is Nothing Then
                lng�д� = 1
                For Each oSubNode1 In Cells
                    Dim lngElementKey As Long
                    lngElementKey = Split(oSubNode1.selectSingleNode("��������").Text, "|")(0)
                    If lngElementKey > 0 Then    'Ҫ�ش���
                        Set oSubNode2 = oNode.selectSingleNode("Elements").selectSingleNode("*[Key=" & oSubNode1.selectSingleNode("Key").Text & " ]")
                        If Not oSubNode2 Is Nothing Then
                            ReDim Preserve ArraySQL(1 To UBound(ArraySQL) + 1) As String
                            ArraySQL(UBound(ArraySQL)) = "Zl_�����ļ��ṹ_Update(" & zlDatabase.GetNextId("�����ļ��ṹ") & "," & lngTempID & "," & IIf(oSubNode2.selectSingleNode("��ID").Text = 0, "NULL", oSubNode2.selectSingleNode("��ID").Text) & "," & _
                            IIf(oSubNode2.selectSingleNode("�������").Text = 0, "NULL", oSubNode2.selectSingleNode("�������").Text) & ",4," & oSubNode2.selectSingleNode("Key").Text & "," & IIf(oSubNode2.selectSingleNode("��������").Text, 1, 0) & ",'" & _
                            oSubNode2.selectSingleNode("��������").Text & "'," & lng�д� & ",'" & Replace(oSubNode2.selectSingleNode("�����ı�").Text, "'", "' || chr(39) || '") & "'," & IIf(oSubNode2.selectSingleNode("�Ƿ���").Text, 1, 0) & ",NULL,NULL,NULL," & _
                            IIf(CheckValid(oSubNode2.selectSingleNode("����Ҫ��ID").Text, oSubNode2.selectSingleNode("Ҫ������").Text), oSubNode2.selectSingleNode("����Ҫ��ID").Text, "NULL") & "," & _
                            oSubNode2.selectSingleNode("�滻��").Text & ",'" & oSubNode2.selectSingleNode("Ҫ������").Text & "'," & oSubNode2.selectSingleNode("Ҫ������").Text & "," & oSubNode2.selectSingleNode("Ҫ�س���").Text & "," & _
                            oSubNode2.selectSingleNode("Ҫ��С��").Text & ",'" & oSubNode2.selectSingleNode("Ҫ�ص�λ").Text & "'," & oSubNode2.selectSingleNode("Ҫ�ر�ʾ").Text & "," & oSubNode2.selectSingleNode("������̬").Text & ",'" & oSubNode2.selectSingleNode("Ҫ��ֵ��").Text & "')"
                        End If
                    Else '�ı�
                        ReDim Preserve ArraySQL(1 To UBound(ArraySQL) + 1) As String
                         ArraySQL(UBound(ArraySQL)) = "Zl_�����ļ��ṹ_Update(" & zlDatabase.GetNextId("�����ļ��ṹ") & "," & lngTempID & "," & oSubNode1.selectSingleNode("��ID").Text & ",NULL," & _
                        "2," & oSubNode1.selectSingleNode("Key").Text & ",NULL,'" & oSubNode1.selectSingleNode("��������").Text & "'," & lng�д� & ",'" & Replace(oSubNode1.selectSingleNode("�����ı�").Text, "'", "' || chr(39) || '") & "')"
                    End If
                    lng�д� = lng�д� + 1
                Next
            End If
            'ͼƬ����
            If Not TablePictures Is Nothing Then
                For Each oSubNode1 In TablePictures
                        ReDim Preserve ArraySQL(1 To UBound(ArraySQL) + 1) As String
                        lngID = zlDatabase.GetNextId("�����ļ��ṹ")
                        ArraySQL(UBound(ArraySQL)) = "Zl_�����ļ��ṹ_Update(" & lngID & "," & lngTempID & "," & IIf(oSubNode1.selectSingleNode("��ID").Text = 0, "NULL", oSubNode1.selectSingleNode("��ID").Text) & "," & _
                        oSubNode1.selectSingleNode("�������").Text & ",5," & oSubNode1.selectSingleNode("Key").Text & "," & IIf(oSubNode1.selectSingleNode("��������").Text, 1, 0) & ",'" & oSubNode1.selectSingleNode("��������").Text & "'," & _
                         lng�д� & ",'" & oSubNode1.selectSingleNode("�����ı�").Text & "'," & IIf(oSubNode1.selectSingleNode("�Ƿ���").Text, 1, 0) & ")"
                        oStream.Type = adTypeBinary
                        oStream.Open
                        oStream.Write oSubNode1.selectSingleNode("OrigPic").nodeTypedValue
                        strPic = App.Path & "\OrigPic" & Timer & ".jpg"
                        oStream.SaveToFile strPic, adSaveCreateOverWrite
                        '-- ���� GDI+ Dll
                        GpInput.GdiplusVersion = 1
                        If (mGdIpEx.GdiplusStartup(m_GDIpToken, GpInput) <> [OK]) Then
                            '����BMP��ʽ���棡������ͼƬ���
                            SavePicture TempPic, strPic       '�����ʽΪBMP��ʽ
                        Else
                            '����JPEGѹ����ʽ����
                            Call oDIB.CreateFromStdPicture(TempPic, DIBPal, DIBDither)
                            'ѹ���洢
                            Call mGdIpEx.SaveDIB(oDIB, strFileName, [ImageJPEG], 100)          '90%��JPEGͼƬѹ������
                        End If
                        'Unload the GDI+ Dll
                        Call mGdIpEx.GdiplusShutdown(m_GDIpToken)
                        gstrSQL = "select ����ID from �����ļ�ͼ�� where ����ID=[1]"
                        Call zlBlobSql(2, lngID, strPic, ArraySQL)
                        oStream.Close
                Next
            End If
        Next
    End If
    '��XML��ȡ����ͼƬ��Ϣ
    If Not oFileRoot.selectSingleNode("Pictures") Is Nothing Then Set Pictures = oFileRoot.selectSingleNode("Pictures").selectNodes("Picture")
    If Not Pictures Is Nothing Then
        For Each oNode In Pictures
            ReDim Preserve ArraySQL(1 To UBound(ArraySQL) + 1) As String
            lngID = zlDatabase.GetNextId("�����ļ��ṹ")
            ArraySQL(UBound(ArraySQL)) = "Zl_�����ļ��ṹ_Update(" & lngID & "," & lngTempID & "," & IIf(oNode.selectSingleNode("��ID").Text = 0, "NULL", oNode.selectSingleNode("��ID").Text) & "," & _
            oNode.selectSingleNode("�������").Text & ",5," & oNode.selectSingleNode("Key").Text & "," & IIf(oNode.selectSingleNode("��������").Text, 1, 0) & ",'" & oNode.selectSingleNode("��������").Text & "'," & _
            "NULL" & ",'" & oNode.selectSingleNode("�����ı�").Text & "'," & IIf(oNode.selectSingleNode("�Ƿ���").Text, 1, 0) & ")"
            oStream.Type = adTypeBinary
            oStream.Open
            oStream.Write oNode.selectSingleNode("OrigPic").nodeTypedValue
            strPic = App.Path & "\OrigPic" & Timer & ".jpg"
            oStream.SaveToFile strPic, adSaveCreateOverWrite
            '-- ���� GDI+ Dll
            GpInput.GdiplusVersion = 1
            If (mGdIpEx.GdiplusStartup(m_GDIpToken, GpInput) <> [OK]) Then
                '����BMP��ʽ���棡������ͼƬ���
                SavePicture TempPic, strPic       '�����ʽΪBMP��ʽ
            Else
                '����JPEGѹ����ʽ����
                Call oDIB.CreateFromStdPicture(TempPic, DIBPal, DIBDither)
                'ѹ���洢
                Call mGdIpEx.SaveDIB(oDIB, strFileName, [ImageJPEG], 100)          '90%��JPEGͼƬѹ������
            End If
            'Unload the GDI+ Dll
            Call mGdIpEx.GdiplusShutdown(m_GDIpToken)
            gstrSQL = "select ����ID from �����ļ�ͼ�� where ����ID=[1]"
            Call zlBlobSql(2, lngID, strPic, ArraySQL)
            oStream.Close
        Next
    End If
    '���ڴ���
     ReDim Preserve ArraySQL(1 To UBound(ArraySQL) + 1) As String
     gstrSQL = "Zl_�����ļ��ṹ_Commit(" & lngTempID & ")"
     ArraySQL(UBound(ArraySQL)) = gstrSQL
    '=========================================================================================
    '����RTFText��Sql
    '=========================================================================================
    If Not oFileRoot.selectSingleNode("Content") Is Nothing Then
        Set ContentNode = oFileRoot.selectSingleNode("Content")
        Me.RTbContext.TextRTF = ContentNode.selectSingleNode("RTF").Text
        If gobjFSO.FileExists(App.Path & "\TMP.rtf") Then gobjFSO.DeleteFile App.Path & "\TMP.rtf", True    '����Ϊ��ʱ�ļ�
        Me.RTbContext.SaveFile App.Path & "\TMP.rtf"
        strTemp = zlFileZip(App.Path & "\TMP.rtf")
        If gobjFSO.FileExists(App.Path & "\TMP.rtf") Then gobjFSO.DeleteFile App.Path & "\TMP.rtf", True
        If gobjFSO.FileExists(strTemp) Then
            Call zlBlobSql(1, lngTempID, strTemp, ArraySQL)
            gobjFSO.DeleteFile strTemp, True      'ɾ����ʱ�ļ�
        End If
    End If
    '#########################################################################################
    '��������
    '=========================================================================================
bb:    If Not BeginTrans(ArraySQL) Then gcnOracle.RollbackTrans: Err.Clear: GoTo errMsg
       ImportFromXml = True
       Exit Function
errHand:
    ImportFromXml = False
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    ImportFromXml = True
    Exit Function
errMsg:
     Result = ZLCommFun.ShowMsgBox("�����ļ�����", strTempName & ",�������ݸ�ʽ����ȷ���ѱ���  ��", "����(&A),����(&O)", Nothing)
     If Result = "����" Then GoTo bb
End Function
'��������ִ��SQL
Private Function BeginTrans(ByVal ArraySQL As Variant) As Boolean
    On Error GoTo errHand
    Dim i As Long
    gcnOracle.BeginTrans
    For i = 1 To UBound(ArraySQL)
        gstrSQL = ArraySQL(i)
        If Trim(gstrSQL) <> "" Then
            Call zlDatabase.ExecuteProcedure(gstrSQL, "cEPRCompends")
        End If
    Next
    gcnOracle.CommitTrans
    BeginTrans = True
    Exit Function
errHand:
    BeginTrans = False
End Function
'################################################################################################################
'## ���ܣ�  �������Ҫ�ص�ԭʼ�����Ƿ���ڣ�����XML����ʱ����֤��
'################################################################################################################
Public Function CheckValid(ByVal ID As Long, ByVal Name As String) As Boolean
    Dim rs As New Recordset
    gstrSQL = "Select ID From ����������Ŀ Where ID = [1] And ������ = [2]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, ID, Name)
    If rs.EOF Then
        CheckValid = False
    Else
        CheckValid = (rs!ID > 0)
    End If
End Function











