VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{FBAFE9A8-8B26-4559-9D12-D70E36A97BE3}#2.1#0"; "zlRichEditor.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmModelExportOrImport 
   Caption         =   "�������������б�"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9990
   Icon            =   "frmModelExportOrImport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   9990
   StartUpPosition =   1  '����������
   Begin zlRichEditor.Editor Editor1 
      Height          =   495
      Left            =   4560
      TabIndex        =   4
      Top             =   3720
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
   End
   Begin VB.ComboBox cboList 
      Height          =   300
      Left            =   4920
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox PicBtn 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   9990
      TabIndex        =   0
      Top             =   6360
      Width           =   9990
      Begin MSComctlLib.ProgressBar progBar 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   3720
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSFlex8Ctl.VSFlexGrid vsgrid 
      Height          =   1860
      Left            =   720
      TabIndex        =   2
      Top             =   2040
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
      BackColor       =   -2147483639
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16761024
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483639
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
   Begin MSComctlLib.ImageList img16 
      Left            =   4440
      Top             =   360
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
            Picture         =   "frmModelExportOrImport.frx":6852
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmModelExportOrImport.frx":6DEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmModelExportOrImport.frx":7386
            Key             =   "ǩ��"
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox RTbFootText 
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmModelExportOrImport.frx":76D8
   End
   Begin RichTextLib.RichTextBox RTbHeadText 
      Height          =   495
      Left            =   2640
      TabIndex        =   6
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmModelExportOrImport.frx":7775
   End
   Begin RichTextLib.RichTextBox RTbContext 
      Height          =   495
      Left            =   1440
      TabIndex        =   7
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmModelExportOrImport.frx":7812
   End
   Begin XtremeCommandBars.ImageManager imgManager 
      Left            =   5280
      Top             =   360
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmModelExportOrImport.frx":78AF
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
Attribute VB_Name = "frmModelExportOrImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
Private Enum mCols
    col_���� = 0
    col_�ļ�ID = 1
    col_�ļ����� = 2
    col_ID = 3
    col_����ID = 4
    col_��� = 5
    col_�������� = 6
    col_���� = 7
    col_���� = 8
    col_���� = 9
    col_˵�� = 11
    col_ͨ�ü� = 10
    col_���� = 12
    col_��Ա = 13
End Enum
'################################################################################################################
'## ���ܣ�  ��ʾ����/���봰��
'## ������  lngType     :��ʾ���ݣ�0-�����б�  1-�����б�
'##         objParent   :������
'################################################################################################################
Public Sub ShowMe(ByVal objParent As Object, ByVal lngType As Long)
    If lngType = 1 Then
        Me.Caption = "�������������б�"
        Me.vsgrid.Tag = "Export"
        Me.PicBtn.Visible = True
        If Not ExportList Then InitVsGrid ("��ʱû�п��Ե��������ݣ�")
        Me.Show 1, objParent
    Else
        Me.Caption = "�������������б�"
        Me.vsgrid.Tag = "Import"
        Me.PicBtn.Visible = True
        If Not ImportList Then Exit Sub
        Me.Show 1, objParent
    End If
End Sub

Private Sub cboList_Click()
    Dim i As Integer, strTempName As String
    With vsgrid
        If .Row < 1 Then Exit Sub
        strTempName = .TextMatrix(.Row + 1, 2)
        .Cell(flexcpData, .Row, 3, .GetNodeRow(.Row, flexNTLastChild), 3) = Me.cboList.ItemData(cboList.ListIndex)
        .Cell(flexcpText, .Row, 2, .GetNodeRow(.Row, flexNTLastChild), 2) = Me.cboList.List(cboList.ListIndex)
        .Cell(flexcpText, .Row, 2, .Row, 5) = Me.cboList.List(cboList.ListIndex) & "(ԭ�����ļ���" & .Cell(flexcpData, .Row, 4) & ")"
        .Cell(flexcpForeColor, .Row, 2, .Row, 5) = vbBlue
    End With
End Sub

Private Sub cboList_GotFocus()
    cboList_Click
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
           Case menu_RemoveRow
                If vsgrid.RowOutlineLevel(vsgrid.Row) = 0 Then MsgBox "���ļ��ڵ��в����Ƴ���", vbInformation, gstrSysName:   Exit Sub
                vsgrid.RemoveItem (vsgrid.Row)
                If vsgrid.Rows = 1 Then InitVsGrid ("���������Ҫ����ķ����ļ� ��")
                Me.Tag = Val(Me.Tag) - 1
           Case menu_Clear
                Call InitVsGrid("���������Ҫ����ķ����ļ� ��")
           Case menu_Export
                StartExportToXMLs
           Case menu_Import
                StartImportFromXML
           Case menu_AddFile
                ImportList
           Case menu_EcheckAll
                CheckAllOrClearAll True
           Case menu_EclearAll
                CheckAllOrClearAll False
           Case menu_Unload
                Unload Me
    End Select
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
           Case menu_Import
                Control.Visible = IIf(vsgrid.Tag = "Import", True, False)
                Control.Enabled = IIf(progBar.Visible Or vsgrid.Cols = 1, False, True)
           Case menu_Export
                Control.Visible = IIf(vsgrid.Tag = "Export", True, False)
                Control.Enabled = IIf(progBar.Visible Or vsgrid.Cols = 1, False, True)
           Case menu_AddFile
                Control.Visible = IIf(vsgrid.Cols = 1 And vsgrid.Tag = "Import", True, False)
           Case menu_EcheckAll, menu_EclearAll
                Control.Enabled = IIf(vsgrid.Cols = 1 Or progBar.Visible, False, True)
           Case menu_Unload
                Control.Enabled = IIf(progBar.Visible, False, True)
    End Select
End Sub

Private Sub Form_Load()
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
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
        Set objControl = .Add(xtpControlButton, menu_Export, "����"): objControl.STYLE = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, menu_Import, "����"): objControl.STYLE = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, menu_AddFile, "���"): objControl.STYLE = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, menu_EcheckAll, "ȫѡ"): objControl.STYLE = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, menu_EclearAll, "ȫ��"): objControl.STYLE = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, menu_Unload, "�˳�"): objControl.STYLE = xtpButtonIconAndCaption
        objControl.STYLE = xtpButtonIconAndCaption  'ͬʱ��ʾͼ�������
    End With
    Me.cbsThis.ActiveMenuBar.Visible = False
End Sub

Private Sub Form_Resize()
    Me.vsgrid.Move 0, 500, Me.ScaleWidth, Me.ScaleHeight - Me.PicBtn.Height - 500
    Me.PicBtn.Move 0, vsgrid.Height + 500, Me.ScaleWidth, PicBtn.Height
    Me.progBar.Move 0, 60, PicBtn.Width, progBar.Height
    If vsgrid.Rows = 1 Or vsgrid.Cols = 1 Then
        vsgrid.ROWHEIGHT(0) = Me.ScaleHeight
    End If
End Sub
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
    vsgrid.ROWHEIGHT(0) = Me.ScaleHeight
End Function
'�����б����
Private Function ExportList() As Boolean
    Dim strType As String, strFileName As String, i As Long, j As Long, k As Long, lngRow As Long
    Dim rsFiles As ADODB.Recordset              '�ļ����ݼ�
    Dim rsModels As New ADODB.Recordset         '�������ݼ�
    On Error GoTo errHand
    gstrSQL = "(Select decode(f.����,1,'1-���ﲡ��',2,'2-סԺ����',4,'4-������',5,'5-����֤������',6,'6-֪���ļ�',7,'���Ƶ���') as ����," & _
              "  f.id as �ļ�ID,f.���� as �ļ����� From ��������Ŀ¼ l, ���ű� d," & _
              "  ��Ա�� p,�����ļ��б� f Where l.����id = d.Id and l.�ļ�id=f.id  And l.��Աid = p.Id and decode(l.����,null,0,0,0)=0 group by f.����,f.id, f.����" & _
              "  Union All" & _
              "  Select decode(f.����,1,'1-���ﲡ��',2,'2-סԺ����',4,'4-������',5,'5-����֤������',6,'6-֪���ļ�',7,'���Ƶ���') as ����," & _
              "  0,null  From ��������Ŀ¼ l, ���ű� d," & _
              "  ��Ա�� p,�����ļ��б� f Where l.����id = d.Id and l.�ļ�id=f.id  And l.��Աid = p.Id and decode(l.����,null,0,0,0)=0 group by ����) order by ����,�ļ�ID"
        Set rsFiles = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
         '��������
        If rsFiles.RecordCount < 1 Then
             Call InitVsGrid("��ʱû�п��Ե����Ĳ����ļ� ��"):  rsFiles.Close
             Exit Function
        End If
        With vsgrid
            .Clear: .FixedRows = 1: .Cols = 14: .ROWHEIGHT(0) = 10
            '���÷���
            .OutlineCol = 0: .OutlineBar = flexOutlineBarCompleteLeaf
            .BackColorSel = vbWhite
            .TextMatrix(0, col_���) = "���": .TextMatrix(0, col_��������) = "����": .TextMatrix(0, col_����) = "����": .TextMatrix(0, col_ͨ�ü�) = "ͨ�ü�": .TextMatrix(0, col_˵��) = "˵��":
            For i = 1 To rsFiles.RecordCount
                '��������ڵ�
                If strType <> rsFiles!���� Then
                    .AddItem ""
                    Me.Tag = Val(Me.Tag) + 1
                    lngRow = Val(Me.Tag)
                    For k = 2 To .Cols - 1
                        .TextMatrix(lngRow, k) = NVL(rsFiles("����").Value)
                        .ColAlignment(k) = flexAlignLeftCenter
                        .ColWidth(k) = 300
                    Next k
                    .Cell(flexcpBackColor, lngRow, 0, lngRow, 13) = &HFFC0C0
                    .IsSubtotal(lngRow) = True
                    .Cell(flexcpData, lngRow, 1) = 1
                    .MergeCells = flexMergeFree
                    .MergeRow(lngRow) = True '�Ƿ������кϲ�
                     strType = rsFiles!����
                Else
                    '�����ļ��ڵ�
                    .AddItem ""
                    Me.Tag = Val(Me.Tag) + 1
                    lngRow = Val(Me.Tag)
                    For k = 3 To .Cols - 1
                        .TextMatrix(lngRow, k) = NVL(rsFiles("�ļ�����").Value)
                    Next k
                    .Cell(flexcpData, lngRow, 2) = NVL(rsFiles("�ļ�ID").Value)
                    .Cell(flexcpBackColor, lngRow, 2, lngRow, 12) = &H80000016
                    .IsSubtotal(lngRow) = True
                    .RowOutlineLevel(lngRow) = 1
                    .MergeCells = flexMergeFree
                    .MergeRow(lngRow) = True '�Ƿ������кϲ�
                    
                    gstrSQL = "Select l.Id,l.�ļ�id,l.���, l.����,l.����,Nvl(l.����, 'δ����') As ����,l.����,l.˵��, l.ͨ�ü�,d.���� As ����," & _
                              "p.���� As ��Ա,Decode(l.����, Null, 1, 2) As ���� From ��������Ŀ¼ l, ���ű� d, ��Ա�� p Where l.����id = d.Id " & _
                              "And l.��Աid = p.Id and l.�ļ�id=" & rsFiles("�ļ�ID").Value & " Order By Decode(l.����, Null, 1, 2), l.����, l.���"
                    Set rsModels = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
                    rsModels.MoveFirst
                    '���ط��Ľڵ�
                    Do While Not rsModels.EOF
                    '0-ȫԺͨ��;1-����ͨ��;2-����ʹ��
                        .AddItem ""
                        Me.Tag = Val(Me.Tag) + 1
                        lngRow = Val(Me.Tag)
                        .Cell(flexcpData, lngRow, col_���) = NVL(rsModels!ID)
                        .TextMatrix(lngRow, col_���) = NVL(rsModels!���)
                        .TextMatrix(lngRow, col_��������) = NVL(rsModels!����)
                        .TextMatrix(lngRow, col_����) = NVL(rsModels!����)
                        .TextMatrix(lngRow, col_����) = NVL(rsModels!����)
                        .TextMatrix(lngRow, col_����) = NVL(rsModels!����)
                        .TextMatrix(lngRow, col_˵��) = NVL(rsModels!˵��)
                        .Cell(flexcpData, lngRow, col_ͨ�ü�) = NVL(rsModels!ͨ�ü�)
                        .Cell(flexcpBackColor, lngRow, 5, lngRow, 13) = &HE0E0E0
                        Select Case Val(rsModels!ͨ�ü�)
                               Case 0
                                .TextMatrix(lngRow, col_ͨ�ü�) = "ȫԺͨ��"
                               Case 1
                               .TextMatrix(lngRow, col_ͨ�ü�) = "����ͨ��"
                               Case 2
                               .TextMatrix(lngRow, col_ͨ�ü�) = "����ʹ��"
                        End Select
                        .TextMatrix(lngRow, col_��Ա) = NVL(rsModels!��Ա)
                        .RowOutlineLevel(lngRow) = 2
                        rsModels.MoveNext
                    Loop
                End If
                rsFiles.MoveNext
           Next i
            .ColWidth(0) = 400: .ColWidth(1) = 270: .ColWidth(2) = 0: .ColWidth(4) = 270
            .ColWidth(col_��������) = 1500: .ColWidth(col_����) = 1000: .ColWidth(col_˵��) = 1000: .ColWidth(col_ͨ�ü�) = 1000: .ColWidth(col_���) = 700
            .ColWidth(col_����) = 0: .ColWidth(col_����) = 0: .ColWidth(col_����) = 0: .ColWidth(col_��Ա) = 0
           '����б���ಿ�ֵı߿���
           For i = 1 To vsgrid.Rows - 1
                If .IsSubtotal(i) = True Then
                    .GetNode(i).Expanded = True
                End If
                If .RowOutlineLevel(i) = 2 Then
                    .Cell(flexcpPicture, i, 4) = img16.ListImages("Check").Picture
                    .Cell(flexcpData, i, 4) = 1
                    .CellBorderRange i, 0, i, 3, vbWhite, 1, 0, 0, 1, 1, 1
                End If
                If .RowOutlineLevel(i) = 1 Then
                    .Cell(flexcpPicture, i, 1) = img16.ListImages("Check").Picture
                    .Cell(flexcpData, i, 1) = 1
                    .CellBorderRange i, 0, i, 13, &H80000016, 1, 1, 1, 1, 1, 1
                End If
                If .RowOutlineLevel(i) = 0 Then
                    .Cell(flexcpPicture, i, 1) = img16.ListImages("Check").Picture
                    .Cell(flexcpData, i, 1) = 1
                    .CellBorderRange i, 0, i, 13, &H80000016, 1, 1, 1, 1, 1, 1
                End If
           Next i
           '----------------------------------------
           vsgrid.RemoveItem (lngRow + 1)
           If vsgrid.Rows > 1 Then vsgrid.Row = 2
    End With
    rsFiles.Close
    ExportList = True
    Exit Function
errHand:
    ExportList = False
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function ImportList() As Boolean
    Dim i As Integer, j As Integer, k As Integer, l As Integer, lngDemoId As Long, lngRow As Long
    Dim strXMLPath As String                'XML·��
    Dim strTempPath As String               '��ʱ·��
    Dim strItems As String                  '������Ϣ
    Dim strFileName As String               '�����ļ�����
    Dim strItemsArr As Variant              '���Ļ�����Ϣ����
    Dim strArrXml As Variant                '�ļ���ַ����
    Dim oDoc As New DOMDocument             'Xml�ĵ�
    Dim cDoc As New cEPRDocument            '�ĵ�����
    Dim oRoot As IXMLDOMElement             '���ڵ�
    Dim oFileList As IXMLDOMNodeList        '�ļ��ڵ㼯��
    Dim oDemoList As IXMLDOMNodeList        '���Ľڵ㼯��
    Dim oSubNode As IXMLDOMElement          '�ӽڵ�
    Dim oSubNode1 As IXMLDOMElement         '�ӽڵ�
    Dim rsTemp As New ADODB.Recordset
    
    On Error Resume Next
    dlgThis.MaxFileSize = 32767
    dlgThis.Filter = "*.ZIP|*.zip"
    dlgThis.DialogTitle = "��"
    dlgThis.CancelError = True
    dlgThis.flags = &H10& Or &H80000
    dlgThis.ShowOpen
    If Err.Number = 32755 Then Err.Clear: ImportList = False: Exit Function
    On Error GoTo errHand
    With vsgrid
            '��������VsGrid
            If Val(Me.Tag) < 1 Then
                .Clear
                .FixedRows = 1: .ExplorerBar = flexExSortShow
                .OutlineCol = 0: .OutlineBar = flexOutlineBarComplete
                .Cols = 6: .ColWidth(0) = 200: .Rows = 1: .ColAlignment(1) = flexAlignLeftCenter: .ROWHEIGHT(0) = Me.cboList.Height: .ColAlignment(0) = flexAlignRightCenter
                .ColWidth(1) = 270: .ColWidth(2) = 1500: .ColWidth(3) = 2500: .ColWidth(4) = 2500: .ColWidth(5) = 6000
                .TextMatrix(0, 1) = "ѡ��": .TextMatrix(0, 2) = "�����ļ�": .TextMatrix(0, 3) = "��������": .TextMatrix(0, 4) = "������λ": .TextMatrix(0, 5) = "�ļ�λ��"
            End If
            '�����ļ�ѡ���������б�����
            gstrSQL = "select ID,���� from �����ļ��б� "
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
            If rsTemp.RecordCount = 0 Then MsgBox "��ǰϵͳ�����ڲ����ļ���������Ӳ����ļ���", vbInformation, gstrSysName: Exit Function
            rsTemp.MoveFirst
            Do While Not rsTemp.EOF
                Me.cboList.AddItem (NVL(rsTemp!����, ""))
                Me.cboList.ItemData(i) = NVL(rsTemp!ID, 0)
                 i = i + 1
                Me.cboList.ListIndex = 0
                rsTemp.MoveNext
            Loop
            '��ʼѭ�����ص��뷶���б�
            strXMLPath = dlgThis.Filename
            Me.cboList.Tag = zlFilesUnZip(strXMLPath)         '��������ļ�·��
            If Me.cboList.Tag = "" Then MsgBox "��ѡ����ļ����ݸ�ʽ����ȷ���ѱ���!", vbInformation, gstrSysName: Exit Function
            oDoc.Load Me.cboList.Tag
            'ɾ����ʱ�ļ�
            gobjFSO.DeleteFile (Me.cboList.Tag)
            '�����·�����ļ��ѱ��������ټ���
            For l = 1 To vsgrid.Rows - 1
                If strXMLPath = Trim(vsgrid.TextMatrix(l, 5)) Then
                     MsgBox strArrXml(i) & ",�Ѿ����򿪣������ظ��� ��", vbInformation, gstrSysName: Exit Function
                End If
            Next l
            '��ȡXML�ļ����ڵ�
            Set oRoot = oDoc.selectSingleNode("EPRDemosList")
            If oRoot Is Nothing Then MsgBox "��ѡ����ļ����ݸ�ʽ����ȷ���ѱ���!", vbInformation, gstrSysName: Exit Function
            '��ȡXML�ļ��в����ļ�����
            Set oFileList = oRoot.selectNodes("/EPRDemosList/Kind/File")
            If oFileList.Item(0) Is Nothing Then MsgBox "��ѡ����ļ����ݸ�ʽ����ȷ���ѱ���!", vbInformation, gstrSysName: Exit Function
            '��ʼѭ�������ļ�����
            For Each oSubNode In oFileList
                '���ļ��ڵ�
                Set oDemoList = oSubNode.selectNodes("Demo")
                Me.Tag = Val(Me.Tag) + 1   '��VsGrid��Rows����
                lngRow = Val(Me.Tag)       'ȡ��VsGrid��Rows
                .AddItem ""
                strFileName = NVL(oSubNode.getAttribute("FileName"))
                gstrSQL = "select ID from �����ļ��б� where ����='" & strFileName & "'"
                '�ж��ļ��Ƿ����
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
                    .Cell(flexcpForeColor, lngRow, 2, lngRow, 5) = IIf(rsTemp.RecordCount > 0, vbBlue, vbMagenta)
                If rsTemp.RecordCount > 0 Then
                   .Cell(flexcpData, lngRow, 3) = Val(rsTemp!ID)                        '���ļ�ID
                   .Cell(flexcpData, lngRow, 1) = 1                                     '����ѡ���е�ֵ
                   .Cell(flexcpPicture, lngRow, 1) = img16.ListImages("Check").Picture  '����ѡ��ͼƬ
                End If
                .Cell(flexcpData, lngRow, 2) = rsTemp.RecordCount  '����ֵ��Ϊ�Ǻ�
                .Cell(flexcpData, lngRow, 4) = strFileName
                For k = 2 To 5
                    .TextMatrix(lngRow, k) = IIf(rsTemp.RecordCount > 0, strFileName, strFileName & "(�ò����ļ��ڵ�ǰ���ݿⲻ���ڣ��뵥���˴�ѡ�����ļ�...)")
                    .ColAlignment(k) = flexAlignLeftCenter
                Next k
                .IsSubtotal(lngRow) = True
                .ROWHEIGHT(lngRow) = Me.cboList.Height
                .MergeCells = flexMergeFree
                .MergeRow(lngRow) = True '�Ƿ������кϲ�
                '�󶨷��Ľڵ�
                For Each oSubNode1 In oDemoList
                     Me.Tag = Val(Me.Tag) + 1
                     lngRow = Val(Me.Tag)
                    .AddItem ""
                    If rsTemp.RecordCount > 0 Then
                        .Cell(flexcpData, lngRow, 1) = 1
                        .Cell(flexcpPicture, lngRow, 1) = img16.ListImages("Check").Picture
                    End If
                     strItems = oSubNode1.getAttribute("Items")
                     lngDemoId = Val(oSubNode1.getAttribute("ID"))
                     strItemsArr = Split(strItems, "|")
                    .ROWHEIGHT(lngRow) = Me.cboList.Height
                    .TextMatrix(lngRow, 2) = strFileName                    '���ļ�����
                    .TextMatrix(lngRow, 3) = strItemsArr(0)                 '�󶨷�������
                    .TextMatrix(lngRow, 4) = oRoot.getAttribute("UnitName") '�󶨵�����λ
                    .TextMatrix(lngRow, 5) = strXMLPath                                                    '���ļ�·��
                    .Cell(flexcpData, lngRow, 3) = .Cell(flexcpData, .GetNodeRow(lngRow, flexNTParent), 3) '����ѡ���е�ֵ
                    .Cell(flexcpForeColor, lngRow, 3) = IIf(rsTemp.RecordCount > 0, vbBlue, vbMagenta)     '����ɫ����
                    .Cell(flexcpData, lngRow, 4) = lngDemoId                '�󶨷���ID
                    .Cell(flexcpData, lngRow, 5) = strItems                 '�󶨷��Ļ�������
                    .RowOutlineLevel(lngRow) = 1
                Next
            Next
    End With
    ImportList = True: Exit Function
errHand:
    ImportList = False
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub Form_Unload(Cancel As Integer)
    If Me.cboList.Tag <> "" Then
        gobjFSO.DeleteFile (gobjFSO.GetParentFolderName(Replace(Me.cboList.Tag, "_�����б�.xml", "_������Ϣ.xml")))
    End If
End Sub

'�����¼�
Private Sub vsgrid_Click()
    If Not vsgrid.MouseIcon Is Nothing And vsgrid.MouseRow > 0 Then
         CheckItems vsgrid.Row
    End If
    If vsgrid.Tag = "Import" Then
        If vsgrid.IsSubtotal(vsgrid.Row) And (vsgrid.MouseCol = 2 Or vsgrid.MouseCol = 3 Or vsgrid.MouseCol = 4) And vsgrid.Cell(flexcpData, vsgrid.Row, 2) <> 1 Then
            Me.cboList.Visible = True
            Me.cboList.Move vsgrid.Cell(flexcpLeft, vsgrid.Row, 2), vsgrid.Cell(flexcpTop, vsgrid.Row, 1) + vsgrid.ROWHEIGHT(vsgrid.Row) * 2 - 100, vsgrid.ColWidth(1) + vsgrid.ColWidth(3)
        Else
            Me.cboList.Visible = False
        End If
    End If
End Sub
'###############################################################
'# ������ ѡ��Vsgrid��ĳ��
'# ������ lngRow :�к�
'###############################################################
Private Sub CheckItems(ByVal lngRow As Long)
    Dim i As Long
    With vsgrid
        If .Tag = "Export" Then
            Select Case .RowOutlineLevel(lngRow)
                   Case 0  'һ��
                     .Cell(flexcpData, lngRow, 1) = IIf(.Cell(flexcpData, lngRow, 1) = 1, 0, 1)
                     .Cell(flexcpPicture, lngRow, 1) = IIf(.Cell(flexcpData, lngRow, 1) = 1, img16.ListImages("Check").Picture, Nothing)
                     For i = lngRow To .GetNodeRow(lngRow, flexNTLastChild)
                        If .RowOutlineLevel(i) = 2 Then
                            .Cell(flexcpData, i, 4) = .Cell(flexcpData, lngRow, 1)
                            .Cell(flexcpPicture, i, 4) = IIf(.Cell(flexcpData, i, 4) = 1, img16.ListImages("Check").Picture, Nothing)
                        ElseIf .RowOutlineLevel(i) = 1 Then
                            .Cell(flexcpData, i, 1) = .Cell(flexcpData, lngRow, 1)
                            .Cell(flexcpPicture, i, 1) = IIf(.Cell(flexcpData, lngRow, 1) = 1, img16.ListImages("Check").Picture, Nothing)
                        End If
                     Next i
                   Case 1 '����
                     .Cell(flexcpData, lngRow, 1) = IIf(.Cell(flexcpData, lngRow, 1) = 1, 0, 1)
                     .Cell(flexcpPicture, lngRow, 1) = IIf(.Cell(flexcpData, lngRow, 1) = 1, img16.ListImages("Check").Picture, Nothing)
                     For i = lngRow To .GetNodeRow(lngRow, flexNTLastChild)
                        If .RowOutlineLevel(i) = 2 Then
                            .Cell(flexcpData, i, 4) = .Cell(flexcpData, lngRow, 1)
                            .Cell(flexcpPicture, i, 4) = IIf(.Cell(flexcpData, i, 4) = 1, img16.ListImages("Check").Picture, Nothing)
                        End If
                     Next i
                   Case 2  '����
                        .Cell(flexcpData, lngRow, 4) = IIf(.Cell(flexcpData, lngRow, 4) = 1, 0, 1)
                        .Cell(flexcpPicture, lngRow, 4) = IIf(.Cell(flexcpData, lngRow, 4) = 1, img16.ListImages("Check").Picture, Nothing)
            End Select
        Else
            If Val(vsgrid.Cell(flexcpData, lngRow, 3)) = 0 Then
                MsgBox "���������ļ������ڣ�����ѡ�����ļ���", vbInformation, gstrSysName: Exit Sub
            End If
            If vsgrid.IsSubtotal(lngRow) Then
              For i = lngRow To .GetNodeRow(lngRow, flexNTLastChild)
                .Cell(flexcpData, i, 1) = IIf(.Cell(flexcpData, lngRow, 1) = 1, 0, 1)
                .Cell(flexcpPicture, i, 1) = IIf(.Cell(flexcpData, lngRow, 1) = 1, img16.ListImages("Check").Picture, Nothing)
                .Cell(flexcpData, i, 1) = IIf(.Cell(flexcpData, lngRow, 1) = 1, 1, 0)
              Next i
            Else
            .Cell(flexcpData, lngRow, 1) = IIf(.Cell(flexcpData, lngRow, 1) = 1, 0, 1)
            .Cell(flexcpPicture, lngRow, 1) = IIf(.Cell(flexcpData, lngRow, 1) = 1, img16.ListImages("Check").Picture, Nothing)
            End If
        End If
    End With
End Sub
'ȫѡ/ȫ��
Private Sub CheckAllOrClearAll(ByVal blnCheck As Boolean)
    Dim i As Long
    If vsgrid.Tag = "Export" Then
        For i = 1 To vsgrid.Rows - 1
            If vsgrid.RowOutlineLevel(i) = 0 Then
                vsgrid.Cell(flexcpData, i, 1) = IIf(blnCheck, 0, 1)
                CheckItems i
            End If
        Next i
    Else
       For i = 1 To vsgrid.Rows - 1
            If vsgrid.RowOutlineLevel(i) = 0 And vsgrid.Cell(flexcpData, i, 3) <> Empty Then
                vsgrid.Cell(flexcpData, i, 1) = IIf(blnCheck, 0, 1)
                vsgrid.Cell(flexcpPicture, i, 1) = IIf(blnCheck, img16.ListImages("Check").Picture, Nothing)
                CheckItems i
           End If
       Next i
    End If
End Sub
'��ȡ���ĵ������ݼ��ַ���
Private Function GetRowsData(ByVal lngRow As Long) As String
    Dim strRowData As String, intModelID As Long
    Dim rsTemp As New ADODB.Recordset                                                                                            '���ַ������ݼ�������
    With vsgrid
            strRowData = strRowData & .Cell(flexcpText, .GetNodeRow(.GetNodeRow(lngRow, flexNTParent), flexNTParent), 2) & "|"   '----��������  0
            strRowData = strRowData & .Cell(flexcpData, .GetNodeRow(lngRow, flexNTParent), 2) & "|"                              '----�ļ�ID    1
            strRowData = strRowData & .TextMatrix(.GetNodeRow(lngRow, flexNTParent), 3) & "|"                                    '----�ļ�����  2
            strRowData = strRowData & .Cell(flexcpData, lngRow, col_���) & "|"                                                  '---- ����ID   3
            strRowData = strRowData & .TextMatrix(lngRow, col_���) & "|"                                                        '---- ���     4
            strRowData = strRowData & .TextMatrix(lngRow, col_��������) & "|"                                                    '---- �������� 5
            strRowData = strRowData & .TextMatrix(lngRow, col_����) & "|"                                                        '---- ����     6
            strRowData = strRowData & .TextMatrix(lngRow, col_����) & "|"                                                        '---- ����     7
            strRowData = strRowData & .TextMatrix(lngRow, col_˵��) & "|"                                                        '---- ˵��     8
            strRowData = strRowData & .Cell(flexcpData, lngRow, col_ͨ�ü�) & "|"                                                '---- ͨ�ü�   9
            strRowData = strRowData & .TextMatrix(lngRow, col_����) & "|"                                                        '---- ����     10
            strRowData = strRowData & glngDeptId & "|"                                                                           '---- ����ID   11
            strRowData = strRowData & glngUserId & "|"                                                                           '---- ����ԱID 12
            intModelID = Val(.Cell(flexcpData, lngRow, col_���))
            gstrSQL = "Select ���� As ������, ���� As ����ֵ From Table(Cast(f_Segment_������('" & intModelID & "') As ZLHIS.t_Dic_Rowset)) where ���� is not null"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
            '���������ȥ��ĩβ��"|"
            If rsTemp.EOF Then strRowData = Mid(strRowData, 1, Len(strRowData) - 1)
            '��ӷ�������ֵ���ַ�����
            Do While Not rsTemp.EOF
                strRowData = strRowData & rsTemp!������ & ":" & rsTemp!����ֵ & ";"
                rsTemp.MoveNext
            Loop
            strRowData = Mid(strRowData, 1, Len(strRowData) - 1)
    End With
    GetRowsData = strRowData
End Function
Private Sub vsgrid_DblClick()
    If vsgrid.MouseRow < 1 Then Exit Sub
    CheckItems (vsgrid.Row)
End Sub

Private Sub vsgrid_KeyDown(KeyCode As Integer, Shift As Integer)
     With vsgrid
        If .IsSubtotal(.Row) Then
            Select Case KeyCode
              Case vbKeyLeft    '��������
                  .GetNode(.Row).Expanded = False
              Case vbKeySpace   '�ո�ѡ��
                   CheckItems .Row
              Case vbKeyRight   '����չ��
                  .GetNode(.Row).Expanded = True
              Case vbKeyReturn  '�س�չ��/����
                .GetNode(.Row).Expanded = Not .GetNode(.Row).Expanded
              Case vbKeyA       'CTRL+A ȫѡ
                If Shift = 2 Then CheckAllOrClearAll (True)
              Case vbKeyZ       'CTRL+Z ȫ��
                If Shift = 2 Then CheckAllOrClearAll (False)
            End Select
        End If
   End With
End Sub

Private Sub vsgrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If vsgrid.Cols = 1 Or vsgrid.Rows = 1 Then Exit Sub
     If Button = 2 Then
        If vsgrid.Tag = "Import" And vsgrid.MouseCol > 0 And vsgrid.MouseRow > 0 Then
                Dim Popup As CommandBar
                Dim objControl As CommandBarControl
                Set Popup = cbsThis.Add("Popup", xtpBarPopup)
                With Popup.Controls
                    .Add xtpControlButton, menu_RemoveRow, "���б����Ƴ�(&D)"
                    .Add xtpControlButton, menu_Clear, "����б�(&C)"
                End With
                Popup.ShowPopup
        End If
      End If
      
End Sub

Private Sub vsgrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intX As Integer, intW As Integer, intX2 As Integer, intW2 As Integer
    If vsgrid.Cols = 1 Then Exit Sub
    If vsgrid.MouseRow = -1 Then
        vsgrid.MousePointer = flexDefault
        Set vsgrid.MouseIcon = Nothing: Exit Sub
    End If
    If vsgrid.Tag = "Import" Then
         If vsgrid.MouseCol = 1 And vsgrid.MouseRow <> 0 Then
                 vsgrid.MousePointer = flexCustom
                 Set vsgrid.MouseIcon = Me.img16.ListImages(1).Picture
         Else
                 vsgrid.MousePointer = flexDefault
                 Set vsgrid.MouseIcon = Nothing
         End If
    Else
        intX = CSng(vsgrid.ColWidth(0)): intW = CSng(vsgrid.ColWidth(0) + vsgrid.ColWidth(1))
        intX2 = vsgrid.ColWidth(0) + vsgrid.ColWidth(1) + vsgrid.ColWidth(2) + vsgrid.ColWidth(3)
        intW2 = intX2 + vsgrid.ColWidth(4)
        If (X > intX And X < intW And vsgrid.Cell(flexcpData, vsgrid.MouseRow, 1) <> "") Or (X > intX2 And X < intW2 And vsgrid.Cell(flexcpData, vsgrid.MouseRow, 4) <> "") Then
         vsgrid.MousePointer = flexCustom
         Set vsgrid.MouseIcon = Me.img16.ListImages(1).Picture
        Else
             vsgrid.MousePointer = flexDefault
             Set vsgrid.MouseIcon = Nothing
        End If
    End If
   
End Sub
''�����ļ�ȫ�����ĵ���
Private Function StartExportToXMLs() As Boolean
    Dim strListPath As String, strInfoPath As String, strPath As String, strRowData As String, strRows As String, strPathZip As String
    Dim i As Long, j As Long, lngRecId As Long, lngDemoId As Long, lngTime As Long, lngRow As Long
    Dim strItemArr As Variant                '���ַ������ݵ�����
    Dim strRowArr As Variant                 'ѡ�е��к�����
    Dim oDocDemosList As New DOMDocument     'Demo�б��ĵ�
    Dim oDocDemosInfo As New DOMDocument     'Demo��Ϣ�ĵ�
    Dim oRootDemosList As IXMLDOMElement     'Demo�б���ڵ�
    Dim oRootInfo As IXMLDOMElement          'Demo��Ϣ���ڵ�
    Dim cEPRDoc As New cEPRDocument          '�ĵ�����
    Dim oKind As IXMLDOMElement              '����ڵ�
    Dim oFile As IXMLDOMElement              '�ļ��ڵ�
    Dim oDemo As IXMLDOMElement              '���Ľڵ�
    Dim oTempNode As IXMLDOMElement          '��ʱ�ڵ�
        '��ͨסԺ����
        On Error Resume Next
        strPath = zl9ComLib.OS.OpenDir(Me.hWnd, "ָ������Ŀ¼")
        If strPath = "" Then Exit Function
        strPathZip = strPath & "\" & zl9ComLib.GetUnitName & "_����.ZIP"
        If gobjFSO.FileExists(strPathZip) Then
            If MsgBox("���ļ��Ѿ����ڣ��Ƿ��滻��", vbOKCancel + vbQuestion, gstrSysName) = vbOK Then
                gobjFSO.DeleteFile (strPathZip)
            Else
             Exit Function
            End If
        End If
        strListPath = strPath & "\" & zl9ComLib.GetUnitName & "_�����б�.xml"
        strInfoPath = strPath & "\" & zl9ComLib.GetUnitName & "_������Ϣ.xml"
        '����Demo��Ϣ���ڵ�
        If oRootDemosList Is Nothing Then
            Set oRootDemosList = oDocDemosList.createElement("EPRDemosList")
            Call oRootDemosList.setAttribute("UnitName", zl9ComLib.GetUnitName)
            Set oDocDemosList.documentElement = oRootDemosList   '����Ϊ���ڵ�
        End If
        '����Demo�б���ڵ�
        If oRootInfo Is Nothing Then
            Set oRootInfo = oDocDemosInfo.createElement("EPRDemosInfo")
            Call oRootInfo.setAttribute("UnitName", zl9ComLib.GetUnitName)
            Set oDocDemosInfo.documentElement = oRootInfo        '����Ϊ���ڵ�
        End If
        On Error GoTo errHand
            EnableControlBar Me, False    '���ô������/С�����رչ���
            lngTime = GetTickCount
            '���㵼�����ĵĸ���
            For i = 0 To vsgrid.Rows - 1
                If Not vsgrid.Cell(flexcpPicture, i, 4) Is Nothing Then
                    strRows = strRows & "," & i
                End If
            Next i
            strRowArr = Split(Mid(strRows, 2, Len(strRows)), ",")
            '��ʼѭ������
            For lngRow = 0 To UBound(strRowArr)
                DoEvents
                i = strRowArr(lngRow)
                If Not vsgrid.Cell(flexcpPicture, i, 4) Is Nothing Then
                    '��ȡ����Demo�ַ�������
                    strRowData = GetRowsData(i)
                    If strRowData = "" Then MsgBox vsgrid.TextMatrix(lngRow, col_��������) & ": ���ݸ�ʽ����ȷ����������� ��", vbInformation, gstrSysName: Exit Function
                    '�������ַ������Ϊ����
                    strItemArr = Split(strRowData, "|")
                    lngDemoId = Val(strItemArr(3))
                    If gobjFSO.FileExists(strListPath) Then
                        If MsgBox("���ļ��Ѿ����ڣ��Ƿ񸲸ǣ�", vbOKCancel + vbQuestion, gstrSysName) = vbCancel Then Exit Function
                    End If
                    Debug.Print vsgrid.TextMatrix(lngRow, col_��������)
                    If Val(strItemArr(7)) <> 2 Then  '�����ڱ���ʽ
                        '��������ڵ�
                        Set oTempNode = oRootDemosList.selectNodes("/EPRDemosList/Kind[@KindName='" & strItemArr(0) & "']")(0)
                        '�ж��Ƿ��Ѿ�����
                        If oTempNode Is Nothing Then
                            Set oKind = CreateNode(1, oRootDemosList, "Kind", NODE_ELEMENT, "")
                            Call oKind.setAttribute("KindName", strItemArr(0))
                        End If
                        '�����ļ��ڵ�
                        Set oTempNode = oRootDemosList.selectNodes("/EPRDemosList/Kind[@KindName='" & strItemArr(0) & "']/File[@FileName='" & strItemArr(2) & "']")(0)
                        If oTempNode Is Nothing Then
                            Set oFile = CreateNode(1, oKind, "File", NODE_ELEMENT, "")
                            Call oFile.setAttribute("FileName", strItemArr(2))
                            Set oTempNode = oFile
                        End If
                        progBar.Visible = True
                        Call ExportDemosToXML(strRowData, lngDemoId, oTempNode, oRootInfo)
                        progBar.Value = IIf(progBar.Value + progBar.Max / (UBound(strRowArr) + 1) > progBar.Max, progBar.Max, progBar.Value + progBar.Max / (UBound(strRowArr) + 1))
                    End If
                End If
            Next lngRow
        oDocDemosList.Save strListPath
        oDocDemosInfo.Save strInfoPath
        'ѹ��XML�ļ�
        Call zlFilesZip(strListPath & "," & strInfoPath, strPathZip)
        MsgBox "������ɣ�", vbOKOnly + vbInformation, gstrSysName
        Unload Me
        StartExportToXMLs = True
        Exit Function
errHand:
    progBar.Value = 0
    progBar.Visible = False
    EnableControlBar Me, True '�ָ��������/С�����رչ���
    StartExportToXMLs = False
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
'��ʼ����XML�ļ�
Private Function StartImportFromXML() As Boolean
    Dim oDoc As New DOMDocument                 'xml�ĵ�
    Dim oRoot  As IXMLDOMElement            '���ڵ�
    Dim oDemoNodeList As IXMLDOMNodeList    '��ʱ�ļ��ڵ㼯��
    Dim oDemoNode As IXMLDOMElement         '���Ľڵ�
    Dim strCheckItems As String, strNumber As String, strRows As String
    Dim i As Long, j As Long, k As Long, lngDemoId As Long, lngFileID As Long, lngRow As Long, lngTime As Long
    Dim strTermsArr As Variant, strItemArr As Variant, strSQLArr As Variant, strRowArr As Variant
    On Error GoTo errHand
    With vsgrid
        '���㵼�뷶�ĵĸ���
        For i = 1 To .Rows - 1
            If Not .Cell(flexcpPicture, i, 1) Is Nothing And .RowOutlineLevel(i) = 1 Then
                strRows = strRows & "," & i
            End If
        Next i
        If strRows = "" Then MsgBox "��ѡ����Ҫ����ķ��ģ�", vbInformation, gstrSysName: Exit Function
        progBar.Visible = True
        EnableControlBar Me, False  '���ô������/С�����رչ���
        oDoc.Load Replace(Me.cboList.Tag, "_�����б�.xml", "_������Ϣ.xml")    '�滻·��Ϊ������Ϣ�ļ�
        gobjFSO.DeleteFolder (gobjFSO.GetParentFolderName(Me.cboList.Tag))     'ɾ����ʱ�ļ���
        Me.cboList.Tag = "" '���.Tag
        '��ȡ�ļ����ڵ�
        Set oRoot = oDoc.selectSingleNode("EPRDemosInfo")
        If oRoot Is Nothing Then MsgBox "���ļ����ݸ�ʽ����ȷ�������ڴ˴�������ļ���", vbInformation, gstrSysName: Exit Function
        '��ȡ���Ľڵ㼯��
        Set oDemoNodeList = oRoot.selectNodes("/EPRDemosInfo/Demo")
        If oDemoNodeList.Length < 1 Then MsgBox "���ļ��������ݿ���Ϊ�գ������ڴ˴�������ļ���", vbInformation, gstrSysName: Exit Function
        '��ʼѭ������
        strRowArr = Split(Mid(strRows, 2, Len(strRows)), ",")
        For lngRow = 0 To UBound(strRowArr)
                DoEvents
                i = strRowArr(lngRow)
                strTermsArr = Split(.Cell(flexcpData, i, 5), "|")
                lngFileID = Val(.Cell(flexcpData, i, 3))                                    '�ļ�ID
                lngDemoId = zlDatabase.GetNextId("��������Ŀ¼")                            '��ȡ����ID
                strNumber = GetMax("��������Ŀ¼", "���", 5, " Where �ļ�id=" & lngFileID) '�������
                gstrSQL = lngDemoId & "," & lngFileID & ",'" & strNumber & "','" & strTermsArr(0) & "','" & strTermsArr(1) & "'," & 0
                gstrSQL = gstrSQL & ",'" & strTermsArr(2) & "'," & strTermsArr(3) & "," & glngDeptId & "," & glngUserId & ",'" & strTermsArr(4) & "'"
                gstrSQL = "Zl_��������Ŀ¼_Insert(" & gstrSQL & ")"
                If strTermsArr(5) <> "0" Then
                    strTermsArr = Split(strTermsArr(5), ";")
                    For j = 0 To UBound(strTermsArr)
                         gstrSQL = "Zl_������������_Edit(' " & lngDemoId & " ','" & Split(strTermsArr(j), ":")(0) & "','" & Split(strTermsArr(j), ":")(1) & "')"
                          Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                    Next j
                End If
                Set oDemoNode = oRoot.selectSingleNode("/EPRDemosInfo/Demo[@ID='" & .Cell(flexcpData, i, 4) & "']")
                ImportDemosFromXML oDemoNode, lngDemoId, strTermsArr(0), .Cell(flexcpData, i, 4), gstrSQL
                k = k + 1
                progBar.Value = IIf(progBar.Value + progBar.Max / (UBound(strRowArr) + 1) > progBar.Max, progBar.Max, progBar.Value + progBar.Max / (UBound(strRowArr) + 1))
        Next lngRow
        Dim strMsg As String
        MsgBox "������ɣ�", vbOKOnly + vbInformation, gstrSysName
        Unload Me
    End With
    StartImportFromXML = True
    Exit Function
errHand:
    progBar.Value = 0
    progBar.Visible = False
    StartImportFromXML = False
    EnableControlBar Me, True  '���ô������/С�����رչ���
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
'################################################################################################################
'## ���ܣ�  �������ļ�����������һ��XML�ĵ���
'##
'## ������  strModel     :   ���������ַ���
'##         lngDemoId    :   ����ID
'##         oFileNode    :   �ļ��ڵ㣬���������ڴ洢�����б�DOC��
'##         oContextNode :   �������ݽڵ㣬���������ڴ洢������Ϣ��DOC��
'## ���أ�  ����ɹ�������Ture�����򷵻�False��
'################################################################################################################
Private Function ExportDemosToXML(ByVal strModel As String, ByVal lngDemoId As Long, ByRef oFileNode As IXMLDOMElement, ByRef oContextNode As IXMLDOMElement) As Boolean
    Dim i As Long, j As Long, k As Long
    Dim oDoc As New DOMDocument
    Dim oDemoRoot As IXMLDOMElement     '���Ľڵ�
    Dim oRootDemo As IXMLDOMElement     '���ڵ�
    Dim oNode As IXMLDOMElement         '���ڵ�
    Dim CompendsoNode As IXMLDOMNode    '��ٽڵ�
    Dim ElementsNode As IXMLDOMNode     'Ҫ�ؽڵ�
    Dim PicturesNode As IXMLDOMNode     'ͼƬ�ڵ�
    Dim DiagnosisesNode As IXMLDOMNode  '���
    Dim TablesNode As IXMLDOMNode       '���ڵ�
    Dim TableCells As IXMLDOMNode       '������ı����Ͻڵ�
    Dim TableElements As IXMLDOMNode    '�����Ҫ�ؼ��Ͻڵ�
    Dim TablePictures As IXMLDOMNode    '�����ͼƬ���Ͻڵ�
    Dim CellNode As IXMLDOMNode         '��Ԫ��ڵ�
    Dim ContentNode As IXMLDOMNode      '���ݽڵ�
    Dim oSubNode1 As IXMLDOMNode        '�ӽڵ�
    Dim oSubNode2 As IXMLDOMNode        '�ڵ�
    Dim oSubNode3 As IXMLDOMNode        '�ڵ�
    Dim oSubNode4 As IXMLDOMNode        '�ڵ�
    Dim oSubNode5 As IXMLDOMNode        '�ڵ�
    Dim oStream As New ADODB.Stream     '������
    Dim strPath As String               '��ʱ�ļ�Ŀ¼
    Dim strPic As String                '��ʱͼƬ�ļ�
    Dim TempPic As New StdPicture, strTempPic As String
    Dim strObjArr As Variant
    Dim strItemArr As Variant, strTermsArr As Variant
    Dim strContextFile As String, strTemp As String
    Dim rs As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    '----------------------------------------------------------------------------
    On Error GoTo errHand:
    strItemArr = Split(strModel, "|")
    strTermsArr = Split(strItemArr(UBound(strItemArr)), ";")
    strPath = IIf(Environ$("tmp") <> vbNullString, Environ$("tmp"), Environ$("temp"))
    '�洢�������б�DOC��
    Set oRootDemo = CreateNode(1, oFileNode, "Demo", NODE_ELEMENT, "")
    Call oRootDemo.setAttribute("EditType", 1)
    Call oRootDemo.setAttribute("ID", lngDemoId)
    Call oRootDemo.setAttribute("Items", strItemArr(5) & "|" & strItemArr(6) & "|" & strItemArr(8) & "|" & strItemArr(9) & "|" & strItemArr(10) & "|" & IIf(UBound(strTermsArr) > 0, strItemArr(UBound(strItemArr)), 0))
    '�洢��������ϢDOC��
    Set oDemoRoot = CreateNode(1, oContextNode, "Demo", NODE_ELEMENT, "")
    Call oDemoRoot.setAttribute("EditType", 1)
    Call oDemoRoot.setAttribute("ID", lngDemoId)
    '��������RTF�ı�
    strContextFile = zlBlobRead(3, lngDemoId)
    If strContextFile <> "" Then
       strTemp = zlFileUnzip(strContextFile)
       Me.RTbContext.LoadFile strTemp
       gobjFSO.DeleteFile strTemp
       gobjFSO.DeleteFile strContextFile, True
    End If
    '��ȡ���Ľṹ
    gstrSQL = "Select Level, ID, �ļ�id, ��id, �������, ��������, ������, ��������, ��������, �����д�, �����ı�, �Ƿ���, Ԥ�����id,�������ID, �������, ʹ��ʱ��," & vbNewLine & _
            "       ����Ҫ��id, �滻��, Ҫ������, Ҫ������, Ҫ�س���, Ҫ��С��, Ҫ�ص�λ, Ҫ�ر�ʾ, ������̬, Ҫ��ֵ��" & vbNewLine & _
            "From (Select ID, �ļ�id, ��id, �������, ��������, ������, ��������, ��������, �����д�, �����ı�, �Ƿ���,Ԥ�����id,�������ID,�������,ʹ��ʱ��," & vbNewLine & _
            "              ����Ҫ��id, �滻��, Ҫ������, Ҫ������, Ҫ�س���, Ҫ��С��, Ҫ�ص�λ, Ҫ�ر�ʾ, ������̬, Ҫ��ֵ��" & vbNewLine & _
            "       From ������������" & vbNewLine & _
            "       Where �ļ�id = [1] And ������� <> 0)" & vbNewLine & _
            "Start With ��id Is Null" & vbNewLine & _
            "Connect By Prior ID = ��id" & vbNewLine & _
            "Order By �������, �����д�"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngDemoId)
    Do While Not rs.EOF
        Select Case NVL(rs("��������"), 2)
            Case 1  '���
                If CompendsoNode Is Nothing Then Set CompendsoNode = CreateNode(1, oDemoRoot, "Compends", NODE_ELEMENT, "")
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
                If TablesNode Is Nothing Then Set TablesNode = CreateNode(1, oDemoRoot, "Tables", NODE_ELEMENT, "")
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
                              "�滻�� , Ҫ������, Ҫ������, Ҫ�س���, Ҫ��С��, Ҫ�ص�λ, Ҫ�ر�ʾ, ������̬, Ҫ��ֵ�� From ������������ " & _
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
                                    strTempPic = zlBlobRead(4, rsTemp!ID)
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
                 If ElementsNode Is Nothing Then Set ElementsNode = CreateNode(1, oDemoRoot, "Elements", NODE_ELEMENT, "")
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
                 If PicturesNode Is Nothing Then Set PicturesNode = CreateNode(1, oDemoRoot, "Pictures", NODE_ELEMENT, "")
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
                    strTempPic = zlBlobRead(4, rs!ID)
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
            Case 7  '���
                If DiagnosisesNode Is Nothing Then Set DiagnosisesNode = CreateNode(1, oDemoRoot, "Diagnosises", NODE_ELEMENT, "")
                Set oSubNode1 = CreateNode(2, DiagnosisesNode, "Diagnosis", NODE_ELEMENT, "")
                CreateNode 3, oSubNode1, "Key", , NVL(rs!������, "")
                CreateNode 3, oSubNode1, "ID", , NVL(rs!ID, 0)
                CreateNode 3, oSubNode1, "�ļ�ID", , NVL(rs!�ļ�ID, 0)
                CreateNode 3, oSubNode1, "��ID", , NVL(rs!��ID, 0)
                CreateNode 3, oSubNode1, "�������", , NVL(rs!�������, 0)
                CreateNode 3, oSubNode1, "����", , NVL(rs!�����ı�, "")
                CreateNode 3, oSubNode1, "��������", , NVL(rs!��������, "")
        End Select
        rs.MoveNext
    Loop
    'RTF�ı�
    Set oNode = CreateNode(1, oDemoRoot, "Content", NODE_ELEMENT, "")
    Set oSubNode1 = CreateNode(2, oNode, "RTF", NODE_ELEMENT, "")
    CreateNode 3, oSubNode1, "RTFText", NODE_CDATA_SECTION, Replace(Me.RTbContext.TextRTF, "]]>", "]] >")
    ExportDemosToXML = True
    Exit Function
errHand:
    ExportDemosToXML = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
'################################################################################################################
'## ���ܣ�  �������뷶��
'## ������  oDemoNode       :���Ľڵ�
'##         lngFileID       :�ļ�ID
'##         strFileName     :�ļ�����
'##         lngOldId        :�ͷ���ID
'##         strSql          :��������SQL���
'## ���أ�  ����ɹ�������Ture�����򷵻�False��
'################################################################################################################
Private Function ImportDemosFromXML(ByVal oDemoNode As IXMLDOMElement, ByVal lngFileID As Long, ByVal strFileName As String, ByVal lngOldId As Long, ByVal strSQL As String) As Boolean
    Dim oNodeList As IXMLDOMNodeList    '�ڵ㼯��
    Dim oNode As IXMLDOMNode            '�ӽڵ�
    Dim oSubNode1 As IXMLDOMNode        '�ӽڵ�
    Dim oSubNode2 As IXMLDOMNode        '�ӽڵ�
    Dim EPRFileInfoNode As IXMLDOMNode  '������Ϣ�ڵ�
    Dim Compends As IXMLDOMNodeList     '��ٽڵ�
    Dim Elements As IXMLDOMNodeList     'Ҫ�ؽڵ�
    Dim Pictures As IXMLDOMNodeList     'ͼƬ�ڵ�
    Dim Diagnosises As IXMLDOMNodeList
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
    Dim strContextFile As String        '��ʱ������
    Dim strArrNames As Variant          '�ļ���������
    Dim ArraySQL() As String            'SQL����
    Dim lngID As Long, lng�д� As Long
    Dim TempPic As New StdPicture, strTempPic As String
    '------------------------------------
    Dim GpInput As GdiplusStartupInput
    Dim m_GDIpToken         As Long         ' ���ڹر� GDI+
    Dim oDIB As New cDIB
    Dim DIBDither As New cDIBDither
    Dim DIBPal As New cDIBPal
    '-------------------------------------------------------------------------
    '��XML��ȡ�����Ϣ
    On Error GoTo errHand
    ReDim ArraySQL(1 To 2) As String
    ArraySQL(1) = strSQL
    If Not oDemoNode.selectSingleNode("Compends") Is Nothing Then Set Compends = oDemoNode.selectSingleNode("Compends").selectNodes("Compend")
    If Not Compends Is Nothing Then
        For Each oNode In Compends
            lngID = zlDatabase.GetNextId("��������Ŀ¼")
            ReDim Preserve ArraySQL(1 To UBound(ArraySQL) + 1) As String
            ArraySQL(UBound(ArraySQL)) = "Zl_������������_Update(" & lngID & "," & lngFileID & "," & IIf(oNode.selectSingleNode("��ID").Text = 0, "NULL", oNode.selectSingleNode("��ID").Text) & "," & _
                oNode.selectSingleNode("�������").Text & ",1," & oNode.selectSingleNode("Key").Text & "," & IIf(oNode.selectSingleNode("��������").Text, 1, 0) & ",'" & oNode.selectSingleNode("˵��").Text & "',NULL,'" & oNode.selectSingleNode("����").Text & "',NULL," & _
                IIf(oNode.selectSingleNode("�������ID").Text = 0, "NULL", oNode.selectSingleNode("�������ID").Text) & "," & IIf(oNode.selectSingleNode("Ԥ�����ID").Text = 0, "NULL", oNode.selectSingleNode("Ԥ�����ID").Text) & "," & IIf(oNode.selectSingleNode("�������").Text, 1, 0) & ",'" & oNode.selectSingleNode("ʹ��ʱ��").Text & "')"
            '�ı��������ID
            Set oNodeList = oDemoNode.selectNodes("/EPRDemosInfo/Demo[@ID='" & lngOldId & "']//��ID[text()=" & oNode.selectSingleNode("ID").Text & "]")
            For Each oSubNode1 In oNodeList
                oSubNode1.Text = lngID
            Next
        Next
    End If
    Debug.Print lngID
    Debug.Print '--------------------------------'
    '��XML��ȡҪ����Ϣ
    If Not oDemoNode.selectSingleNode("Elements") Is Nothing Then Set Elements = oDemoNode.selectSingleNode("Elements").selectNodes("Element")
    If Not Elements Is Nothing Then
        For Each oNode In Elements
            ReDim Preserve ArraySQL(1 To UBound(ArraySQL) + 1) As String
            ArraySQL(UBound(ArraySQL)) = "Zl_������������_Update(" & zlDatabase.GetNextId("��������Ŀ¼") & "," & lngFileID & "," & IIf(oNode.selectSingleNode("��ID").Text = 0, "NULL", oNode.selectSingleNode("��ID").Text) & "," & _
                oNode.selectSingleNode("�������").Text & ",4," & oNode.selectSingleNode("Key").Text & "," & IIf(oNode.selectSingleNode("��������").Text, 1, 0) & ",'" & oNode.selectSingleNode("��������").Text & "',NULL,'" & _
                Replace(oNode.selectSingleNode("�����ı�").Text, "'", "' || chr(39) || '") & "'," & IIf(oNode.selectSingleNode("�Ƿ���").Text, 1, 0) & ",NULL,NULL,NULL," & _
                 "NULL," & IIf(CheckValid(oNode.selectSingleNode("����Ҫ��ID").Text, oNode.selectSingleNode("Ҫ������").Text), oNode.selectSingleNode("����Ҫ��ID").Text, "NULL") & "," & _
                oNode.selectSingleNode("�滻��").Text & ",'" & oNode.selectSingleNode("Ҫ������").Text & "'," & oNode.selectSingleNode("Ҫ������").Text & "," & oNode.selectSingleNode("Ҫ�س���").Text & "," & _
                oNode.selectSingleNode("Ҫ��С��").Text & ",'" & oNode.selectSingleNode("Ҫ�ص�λ").Text & "'," & oNode.selectSingleNode("Ҫ�ر�ʾ").Text & "," & oNode.selectSingleNode("������̬").Text & ",'" & oNode.selectSingleNode("Ҫ��ֵ��").Text & "')"
        Next
    End If
    '��XML��ȡ�����Ϣ
    If Not oDemoNode.selectSingleNode("Tables") Is Nothing Then Set Tables = oDemoNode.selectSingleNode("Tables").selectNodes("Table")
    If Not Tables Is Nothing Then
        For Each oNode In Tables
            lngID = zlDatabase.GetNextId("��������Ŀ¼")
            '������ṹSQL���
            ReDim Preserve ArraySQL(1 To UBound(ArraySQL) + 1) As String
            ArraySQL(UBound(ArraySQL)) = "Zl_������������_Update(" & lngID & "," & lngFileID & "," & IIf(oNode.selectSingleNode("��ID").Text = 0, "NULL", oNode.selectSingleNode("��ID").Text) & "," & _
            oNode.selectSingleNode("�������").Text & ",3," & oNode.selectSingleNode("Key").Text & "," & IIf(oNode.selectSingleNode("��������").Text, 1, 0) & ",'" & oNode.selectSingleNode("��������").Text & "',NULL,'" & "" & "'," & IIf(oNode.selectSingleNode("�Ƿ���").Text, 1, 0) & _
            "," & IIf(oNode.selectSingleNode("Ԥ�����ID").Text = 0, "NULL", oNode.selectSingleNode("Ԥ�����ID").Text) & ")"
            '������������ĸ�ID
            For Each oSubNode1 In oNode.selectNodes("/EPRDemosInfo/Demo[@ID='" & lngOldId & "']//��ID[text()=" & oNode.selectSingleNode("ID").Text & "]")
                oSubNode1.Text = lngID
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
                            ArraySQL(UBound(ArraySQL)) = "Zl_������������_Update(" & zlDatabase.GetNextId("��������Ŀ¼") & "," & lngFileID & "," & IIf(oSubNode2.selectSingleNode("��ID").Text = 0, "NULL", oSubNode2.selectSingleNode("��ID").Text) & "," & _
                            IIf(oSubNode2.selectSingleNode("�������").Text = 0, "NULL", oSubNode2.selectSingleNode("�������").Text) & ",4," & oSubNode2.selectSingleNode("Key").Text & "," & IIf(oSubNode2.selectSingleNode("��������").Text, 1, 0) & ",'" & _
                            oSubNode2.selectSingleNode("��������").Text & "'," & lng�д� & ",'" & Replace(oSubNode2.selectSingleNode("�����ı�").Text, "'", "' || chr(39) || '") & "'," & IIf(oSubNode2.selectSingleNode("�Ƿ���").Text, 1, 0) & ",NULL,NULL,NULL," & _
                             "NULL," & IIf(CheckValid(oSubNode2.selectSingleNode("����Ҫ��ID").Text, oSubNode2.selectSingleNode("Ҫ������").Text), oSubNode2.selectSingleNode("����Ҫ��ID").Text, "NULL") & "," & _
                            oSubNode2.selectSingleNode("�滻��").Text & ",'" & oSubNode2.selectSingleNode("Ҫ������").Text & "'," & oSubNode2.selectSingleNode("Ҫ������").Text & "," & oSubNode2.selectSingleNode("Ҫ�س���").Text & "," & _
                            oSubNode2.selectSingleNode("Ҫ��С��").Text & ",'" & oSubNode2.selectSingleNode("Ҫ�ص�λ").Text & "'," & oSubNode2.selectSingleNode("Ҫ�ر�ʾ").Text & "," & oSubNode2.selectSingleNode("������̬").Text & ",'" & oSubNode2.selectSingleNode("Ҫ��ֵ��").Text & "')"
                        End If
                    Else '�ı�
                        ReDim Preserve ArraySQL(1 To UBound(ArraySQL) + 1) As String
                         ArraySQL(UBound(ArraySQL)) = "Zl_������������_Update(" & zlDatabase.GetNextId("��������Ŀ¼") & "," & lngFileID & "," & oSubNode1.selectSingleNode("��ID").Text & ",NULL," & _
                        "2," & oSubNode1.selectSingleNode("Key").Text & ",NULL,'" & oSubNode1.selectSingleNode("��������").Text & "'," & lng�д� & ",'" & Replace(oSubNode1.selectSingleNode("�����ı�").Text, "'", "' || chr(39) || '") & "')"
                    End If
                    lng�д� = lng�д� + 1
                Next
            End If
            'ͼƬ����
            If Not TablePictures Is Nothing Then
                For Each oSubNode1 In TablePictures
                        ReDim Preserve ArraySQL(1 To UBound(ArraySQL) + 1) As String
                        lngID = zlDatabase.GetNextId("��������Ŀ¼")
                        ArraySQL(UBound(ArraySQL)) = "Zl_������������_Update(" & lngID & "," & lngFileID & "," & IIf(oSubNode1.selectSingleNode("��ID").Text = 0, "NULL", oSubNode1.selectSingleNode("��ID").Text) & "," & _
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
                        gstrSQL = "select ����ID from ��������ͼ�� where ����ID=[1]"
                        Call zlBlobSql(4, lngID, strPic, ArraySQL)
                        oStream.Close
                Next
            End If
        Next
    End If
    '��XML��ȡ�����Ϣ
     If Not oDemoNode.selectSingleNode("Diagnosises") Is Nothing Then Set Diagnosises = oDemoNode.selectSingleNode("Diagnosises").selectNodes("Diagnosis")
     If Not Diagnosises Is Nothing Then
        For Each oNode In Diagnosises
            ReDim Preserve ArraySQL(1 To UBound(ArraySQL) + 1) As String
            lngID = zlDatabase.GetNextId("��������Ŀ¼")
            ArraySQL(UBound(ArraySQL)) = "Zl_������������_Update(" & lngID & "," & lngFileID & "," & _
            IIf(oNode.selectSingleNode("��ID").Text = 0, "NULL", oNode.selectSingleNode("��ID").Text) & "," & oNode.selectSingleNode("�������").Text & ",7," & _
            oNode.selectSingleNode("Key").Text & ",0,'" & oNode.selectSingleNode("��������").Text & "',NULL,'" & oNode.selectSingleNode("����").Text & "')"
        Next
    End If
    '��XML��ȡ����ͼƬ��Ϣ
    If Not oDemoNode.selectSingleNode("Pictures") Is Nothing Then Set Pictures = oDemoNode.selectSingleNode("Pictures").selectNodes("Picture")
    If Not Pictures Is Nothing Then
        For Each oNode In Pictures
            ReDim Preserve ArraySQL(1 To UBound(ArraySQL) + 1) As String
            lngID = zlDatabase.GetNextId("��������Ŀ¼")
            ArraySQL(UBound(ArraySQL)) = "Zl_������������_Update(" & lngID & "," & lngFileID & "," & IIf(oNode.selectSingleNode("��ID").Text = 0, "NULL", oNode.selectSingleNode("��ID").Text) & "," & _
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
            gstrSQL = "select ����ID from ��������ͼ�� where ����ID=[1]"
            Call zlBlobSql(4, lngID, strPic, ArraySQL)
            oStream.Close
        Next
    End If
    '���ڴ���
     ReDim Preserve ArraySQL(1 To UBound(ArraySQL) + 1) As String
     gstrSQL = "zl_������������_commit(" & lngFileID & ")"
     ArraySQL(UBound(ArraySQL)) = gstrSQL
    '=========================================================================================
    '����RTFText��Sql
    '=========================================================================================
    If Not oDemoNode.selectSingleNode("Content") Is Nothing Then
        Set ContentNode = oDemoNode.selectSingleNode("Content")
        Me.RTbContext.TextRTF = ContentNode.selectSingleNode("RTF").Text
        If gobjFSO.FileExists(App.Path & "\TMP.rtf") Then gobjFSO.DeleteFile App.Path & "\TMP.rtf", True    '����Ϊ��ʱ�ļ�
        Me.RTbContext.SaveFile App.Path & "\TMP.rtf"
        strTemp = zlFileZip(App.Path & "\TMP.rtf")
        If gobjFSO.FileExists(App.Path & "\TMP.rtf") Then gobjFSO.DeleteFile App.Path & "\TMP.rtf", True
        If gobjFSO.FileExists(strTemp) Then
            Call zlBlobSql(3, lngFileID, strTemp, ArraySQL)
            gobjFSO.DeleteFile strTemp, True      'ɾ����ʱ�ļ�
        End If
    End If
bb: If Not BeginTrans(ArraySQL) Then gcnOracle.RollbackTrans: Err.Clear: GoTo errHand
    ImportDemosFromXML = True: Exit Function
errHand:
    ImportDemosFromXML = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
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
