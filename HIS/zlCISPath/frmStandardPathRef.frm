VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmStandardPathRef 
   Caption         =   "��׼·���ο�"
   ClientHeight    =   8730
   ClientLeft      =   6345
   ClientTop       =   2085
   ClientWidth     =   13755
   Icon            =   "frmStandardPathRef.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8730
   ScaleWidth      =   13755
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picSTPathList 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8010
      Left            =   0
      ScaleHeight     =   8010
      ScaleWidth      =   4005
      TabIndex        =   10
      Top             =   0
      Width           =   4005
      Begin XtremeReportControl.ReportControl rptStPath 
         Height          =   1695
         Left            =   240
         TabIndex        =   11
         Top             =   5760
         Width           =   2820
         _Version        =   589884
         _ExtentX        =   4974
         _ExtentY        =   2990
         _StockProps     =   0
         SkipGroupsFocus =   0   'False
      End
      Begin VB.PictureBox picFind 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   3855
         TabIndex        =   13
         Top             =   120
         Width           =   3855
         Begin VB.CommandButton cmdImport 
            Caption         =   "����·��"
            Height          =   300
            Left            =   2710
            TabIndex        =   15
            Top             =   360
            Width           =   1100
         End
         Begin VB.TextBox txtInput 
            Height          =   300
            Left            =   435
            MaxLength       =   100
            TabIndex        =   14
            Top             =   0
            Width           =   3375
         End
         Begin VB.Label lblFind 
            BackColor       =   &H00FFFFFF&
            Caption         =   "����"
            Height          =   255
            Left            =   0
            TabIndex        =   16
            Top             =   30
            Width           =   375
         End
      End
      Begin XtremeSuiteControls.TabControl tbcPathName 
         Height          =   1455
         Left            =   120
         TabIndex        =   12
         Top             =   4200
         Width           =   3015
         _Version        =   589884
         _ExtentX        =   5318
         _ExtentY        =   2566
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picStPathDetial 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7335
      Left            =   5280
      ScaleHeight     =   7335
      ScaleWidth      =   7575
      TabIndex        =   1
      Top             =   0
      Width           =   7575
      Begin VB.PictureBox picPathTable 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   240
         ScaleHeight     =   2295
         ScaleWidth      =   6255
         TabIndex        =   3
         Top             =   1800
         Width           =   6255
         Begin VB.Frame fraSplitNS 
            BackColor       =   &H00F0F4E4&
            BorderStyle     =   0  'None
            Height          =   100
            Left            =   0
            TabIndex        =   9
            Top             =   1200
            Width           =   6255
         End
         Begin VSFlex8Ctl.VSFlexGrid vsPathTable 
            Height          =   975
            Left            =   0
            TabIndex        =   4
            Top             =   1320
            Width           =   3585
            _cx             =   6324
            _cy             =   1720
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
            BackColor       =   16777215
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   16777215
            BackColorAlternate=   16777215
            GridColor       =   32768
            GridColorFixed  =   32768
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   3
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   3
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   4
            Cols            =   4
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   20
            RowHeightMax    =   5000
            ColWidthMin     =   20
            ColWidthMax     =   9000
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmStandardPathRef.frx":058A
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
         Begin VB.Frame fra��ͷ 
            BackColor       =   &H00F0F4E4&
            BorderStyle     =   0  'None
            Height          =   1095
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   6255
            Begin VB.Label lbl��ͷ 
               AutoSize        =   -1  'True
               BackColor       =   &H00F0F4E4&
               Height          =   180
               Left            =   120
               TabIndex        =   6
               Top             =   0
               Width           =   90
            End
         End
      End
      Begin XtremeSuiteControls.TabControl tbcStPath 
         Height          =   7335
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   6615
         _Version        =   589884
         _ExtentX        =   11668
         _ExtentY        =   12938
         _StockProps     =   64
      End
      Begin VB.PictureBox picPathCourse 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   0
         ScaleHeight     =   3735
         ScaleWidth      =   6255
         TabIndex        =   7
         Top             =   480
         Width           =   6255
         Begin RichTextLib.RichTextBox rtfPathCourse 
            Height          =   4095
            Left            =   120
            TabIndex        =   8
            Top             =   0
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   7223
            _Version        =   393217
            BackColor       =   16777215
            BorderStyle     =   0
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"frmStandardPathRef.frx":05F6
         End
      End
   End
   Begin VB.Frame fraSplit 
      Caption         =   "Frame1"
      Height          =   7335
      Left            =   5200
      MousePointer    =   9  'Size W E
      TabIndex        =   0
      Top             =   0
      Width           =   45
   End
End
Attribute VB_Name = "frmStandardPathRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrs��         As New ADODB.Recordset      '��׼·����
Private mrs��ͷ��Ϣ     As New ADODB.Recordset      '��׼·�����ı�ͷ�Լ�������Ϣ����Ϣ
Private mlngStPathID    As Long                     'ѡ�еı�׼·����ID
Private mlngFunc        As Long                     '0-��׼·�����鿴,1-����׼·���������ٴ�·����
Private mblnOK          As Boolean                  'mblnFunc=1:����ɹ� ����True
Private mintMode        As Integer                  '0-סԺ��1-����
Private Const M_INT_STEPNUM = 3

Private Enum PathListCols
    COL_ID = 0
    COL_�������� = 1
    COL_���� = 2
    COL_·������ = 3
    COL_�汾˵�� = 4
    COL_�������� = 5
End Enum

Private Enum CATE_TYPE
    IX_·������ = 0
    IX_�������� = 1
End Enum

Public Function ShowMe(frmMain As Object, ByVal lngStPathID As Long, Optional ByVal lngFunc As Long = 0, Optional ByVal intMode As Integer)
'������lngStPathID ѡ�еı�׼·��
    mblnOK = False
    mlngStPathID = lngStPathID
    mlngFunc = lngFunc
    mintMode = intMode
    Me.Show 1, frmMain
    
    ShowMe = mblnOK
End Function

Private Sub FuncFindSTPath(Optional ByVal blnNext As Boolean)
'����:���������ı�����,��λ·����
    Dim i As Long
    Dim blnHave As Boolean
    Dim blnReStart As Boolean
    Dim objRow As Object
    
    Call zlControl.TxtSelAll(txtInput)
    '��ʼ������
    If rptStPath.SelectedRows.count > 0 Then blnHave = True
    If Not blnNext Or blnReStart Or Not blnHave Then
        i = 0    'ReportControl����������0��ʼ
    Else
        i = rptStPath.SelectedRows(0).Index + 1
    End If

    For Each objRow In rptStPath.Rows
        objRow.Expanded = True
    Next
    '����·��
    For i = i To rptStPath.Rows.count - 1
        With rptStPath.Rows(i)
            If .Record.Tag <> "" Then
                If zlStr.IsCharChinese(Trim(txtInput.Text)) Then
                    If .Record(COL_·������).Value Like "*" & Trim(txtInput.Text) & "*" Then
                        Exit For
                    End If
                Else '��������
                    If .Record(COL_��������).Value Like "*" & UCase(Trim(txtInput.Text)) & "*" Then
                        Exit For
                    End If
                End If
        
            End If
        End With
    Next
    
    If i <= rptStPath.Rows.count - 1 Then
        blnReStart = False
        '����ѡ������ʾ�ڿɼ�����,������SelectionChanged�¼�
        Set rptStPath.FocusedRow = rptStPath.Rows(i)

        If rptStPath.Visible Then rptStPath.SetFocus
    Else
        blnReStart = True
        MsgBox IIf(blnNext, "������", "") & "�Ҳ������������ı�׼·����", vbInformation, gstrSysName
    End If
End Sub

Private Sub cmdImport_Click()
'����:����׼·���������ٴ�·����
    Dim strSql As String

    On Error GoTo errH
    If mlngStPathID <> 0 Then
        If mintMode = 1 Then
            strSql = "zl_����·������_Import(" & mlngStPathID & ")"
        Else
            strSql = "zl_�ٴ�·������_Import(" & mlngStPathID & ")"
        End If
        zlDatabase.ExecuteProcedure strSql, Me.Caption
        MsgBox "����ɹ�!", vbInformation + vbOKOnly, gstrSysName
        mblnOK = True
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'mlngFunc=1 ֧�ֲ��ҹ���
    If mlngFunc = 1 Then
        If KeyCode = vbKeyF And Shift = vbCtrlMask Then
            txtInput.SetFocus
            If Trim(txtInput.Text) <> "" Then
                Call FuncFindSTPath
            End If
        ElseIf KeyCode = vbKeyF3 Then
            If Trim(txtInput.Text) <> "" Then
                FuncFindSTPath (True)
            End If
            txtInput.SetFocus
        End If
    End If
End Sub

Private Sub Form_Load()
'���ܣ���ʼ��tbcControl ��reportControl
    Dim objCol        As ReportColumn

    If mlngFunc = 0 Then
        If mintMode = 1 Then
            Me.Caption = "��׼����·���ο�"
        Else
            Me.Caption = "��׼·���ο�"
        End If
        picFind.Visible = False
    Else
        If mintMode = 1 Then
            Me.Caption = "�����׼����·��"
        Else
            Me.Caption = "�����׼·��"
        End If
        picFind.Visible = True
    End If
    'tbcPathName·���ο�
    With Me.tbcPathName
        With .PaintManager
            .Appearance = xtpTabAppearanceExcel
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        
        .InsertItem 0, "��ҽ�ο�", rptStPath.Hwnd, 0
        .InsertItem 1, "��ҽ�ο�", rptStPath.Hwnd, 0
        
        .Item(1).Selected = True
        .Item(0).Selected = True
        
    End With
    
    With rptStPath
        '��ʼ��Report�ؼ�����������
        Set objCol = .Columns.Add(PathListCols.COL_ID, "ID", 20, False)
            objCol.Alignment = xtpAlignmentCenter: objCol.Resizable = True: objCol.AllowDrag = False: objCol.Visible = False
        Set objCol = .Columns.Add(PathListCols.COL_��������, "��������", 80, False)
            objCol.Resizable = True: objCol.Alignment = xtpAlignmentLeft: objCol.AllowDrag = False: objCol.TreeColumn = True: objCol.Groupable = True
        Set objCol = .Columns.Add(PathListCols.COL_����, "����", 50, False)
            objCol.Alignment = xtpAlignmentLeft: objCol.Resizable = True: objCol.AllowDrag = False
        Set objCol = .Columns.Add(PathListCols.COL_·������, "·������", 200, False)
            objCol.Alignment = xtpAlignmentLeft: objCol.Resizable = True: objCol.AllowDrag = False
        Set objCol = .Columns.Add(PathListCols.COL_�汾˵��, "�汾˵��", 70, False)
            objCol.Alignment = xtpAlignmentLeft: objCol.Resizable = True: objCol.AllowDrag = False
        Set objCol = .Columns.Add(PathListCols.COL_��������, "��������", 200, False)
            objCol.Alignment = xtpAlignmentLeft: objCol.Resizable = True: objCol.AllowDrag = False
            
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoItemsText = "û�п���ʾ����Ŀ."
        End With
        
        .AutoColumnSizing = False
        .AllowColumnRemove = False
        .ShowGroupBox = False
        .ShowItemsInGroups = False
        .PreviewMode = True
        .MultipleSelection = False '������SelectionChanged�¼�
    End With
    
    '��ʼ��tbcControl
    With tbcStPath
        With .PaintManager
            .Appearance = xtpTabAppearanceVisualStudio
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
        End With
        
        .AllowReorder = False
        '���μ�������ֻ����ѡ��Լ���׼����
        If mintMode = 1 Then
            Call .InsertItem(0, "��׼��������", picPathCourse.Hwnd, 0)
        Else
            Call .InsertItem(0, "��׼סԺ����", picPathCourse.Hwnd, 0)
        End If
        .Item(0).Selected = True 'Ĭ��ѡ���׼����
    End With
    
    '���ر�׼·��Ŀ¼
    Call LoadStPathList(0, True)
    '����ѡ��ı�׼·��ID����·�����̣�·����������ͷ
    Call LoadPathByID(mlngStPathID, True, 0)
End Sub

Private Sub Form_Resize()
'���ܣ�����tbcPathName��picStPathDetial��λ�ô�С
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.WindowState <> vbMaximized And Me.WindowState <> vbMinimized Then
        Me.Height = IIf(Me.Height < 9000, 9000, Me.Height)
        Me.Width = IIf(Me.Width < 12000, 12000, Me.Width)
    End If

    picSTPathList.Move 0, 0, Me.ScaleWidth * 0.3, Me.ScaleHeight
   
    fraSplit.Left = picSTPathList.Left + picSTPathList.Width + 30
    fraSplit.Height = Me.ScaleHeight
    
    picStPathDetial.Left = fraSplit.Left + fraSplit.Width + 30
    picStPathDetial.Width = Me.ScaleWidth - picStPathDetial.ScaleLeft
    picStPathDetial.Height = Me.ScaleHeight - picStPathDetial.ScaleTop
 
End Sub

Private Sub fraSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'���ܣ�ʵ�ֱ�׼·���嵥���׼·�����������϶���С
    If Button = 1 Then
        If picSTPathList.Width + X > 11000 Or picSTPathList.Width + X < 2000 Then Exit Sub
        
        fraSplit.Left = fraSplit.Left + X
        picSTPathList.Width = fraSplit.Left - 30 - picSTPathList.Left
        picStPathDetial.Left = fraSplit.Left + 30
        picStPathDetial.Width = Me.ScaleWidth - picStPathDetial.Left
        
        Me.Refresh
    End If
    
End Sub

Private Sub picPathCourse_Resize()
'���ܣ�ʵ�ֱ�׼·���������ݵĴ�С����

    rtfPathCourse.Width = picPathCourse.Width - rtfPathCourse.Left - 120
    rtfPathCourse.Height = picPathCourse.Height - rtfPathCourse.Top
    
End Sub

Private Sub picPathTable_Resize()
'���ܣ����ñ���ͷ������ڿؼ���λ�����С

    fra��ͷ.Height = lbl��ͷ.Height + 60
    fra��ͷ.Width = picPathTable.Width
    lbl��ͷ.Width = fra��ͷ.Width
    lbl��ͷ.Width = fra��ͷ.Width - lbl��ͷ.Left
    
    fraSplitNS.Top = fra��ͷ.Top + fra��ͷ.Height
    fraSplitNS.Width = picPathTable.Width
    
    vsPathTable.Top = fraSplitNS.Top + fraSplitNS.Height
    vsPathTable.Height = picPathTable.Height - vsPathTable.Top - 120
    vsPathTable.Width = picPathTable.Width - vsPathTable.Left - 120
    
End Sub

Private Sub picStPathDetial_Resize()
'���ܣ���׼·���������Ĵ�С����

    tbcStPath.Width = picStPathDetial.Width
    tbcStPath.Height = picStPathDetial.Height
    picPathTable.Width = tbcStPath.Width '����picStPathDetial_Resize
    picPathCourse.Width = tbcStPath.Width
    
End Sub

Private Sub picSTPathList_Resize()
    On Error Resume Next
    If mlngFunc = 0 Then
        tbcPathName.Top = 0
        tbcPathName.Left = 0
    Else
        picFind.Move 120, 120, picSTPathList.Width - 240, 850
        tbcPathName.Top = picFind.Height + picFind.Top
        tbcPathName.Left = 0
    End If
    
    tbcPathName.Width = picSTPathList.Width
    tbcPathName.Height = picSTPathList.Height - tbcPathName.Top

End Sub

Private Sub rptStPath_SelectionChanged()
'���ܣ�����ѡ���·��ID,������ID���ر�׼·�������Լ���

    If Me.Visible Then
        If mlngStPathID <> Val(rptStPath.SelectedRows(0).Record.Tag) And Val(rptStPath.SelectedRows(0).Record.Tag) <> 0 Then
            mlngStPathID = Val(rptStPath.SelectedRows(0).Record.Tag)
            tbcPathName.Item(tbcPathName.Selected.Index).Tag = mlngStPathID
            Call LoadPathByID(mlngStPathID, True, 0)
        End If
        
        cmdImport.Enabled = (mlngStPathID <> 0 And mlngFunc = 1)
    End If
    
End Sub

Private Sub tbcPathName_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'����:����ѡ�ѡ���������
    If Me.Visible Then
        mlngStPathID = 0
        Call LoadStPathList(Item.Index)
        Call LoadPathByID(mlngStPathID, True, 0)
    End If
End Sub

Private Sub tbcStPath_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'���ܣ�ѡ���ʱ���ر�����

    If Me.Visible Then
        Call LoadPathByID(mlngStPathID, False, Item.Index)
        picPathCourse.Visible = Item.Index = 0
        picPathTable.Visible = Item.Index <> 0
    End If
End Sub

Private Sub LoadStPathList(ByVal lngIndex As Long, Optional ByVal blnFirst As Boolean)
'���ܣ����ر�׼·��Ŀ¼
'����:lngIndex 0-��ҽ�ο�,1-��ҽ�ο�
'     blnFirst True-�״μ���,False-���״μ���
    Dim objRecord     As ReportRecord
    Dim objPreRecord     As ReportRecord
    Dim objItem       As ReportRecordItem
    Dim i As Long, strDept As String
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
    
On Error GoTo errH
    '��ռ�¼,���������ظ�
    rptStPath.Records.DeleteAll
    
    If blnFirst And mlngStPathID <> 0 Then
        If mintMode = 1 Then
            strSql = "Select Nvl(t.���, 0) as ��� From ��׼����·��Ŀ¼ T Where t.Id = [1]"
        Else
            strSql = "Select Nvl(t.���, 0) as ��� From ��׼·��Ŀ¼ T Where t.Id = [1]"
        End If
        
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngStPathID)
        lngIndex = Val(rsTemp!��� & "")
        tbcPathName.Item(lngIndex).Selected = True 'ѡ��ָ����
    End If

    If mintMode = 1 Then
        strSql = " Select a.Id, a.��������, a.����, a.·������, a.�汾˵��, b.��������" & vbNewLine & _
                 " From ��׼����·��Ŀ¼ A, ��׼����·������ B" & vbNewLine & _
                 " Where a.Id = b.��׼·��id  and Nvl(a.���,0)=[1] order by ��������,ID "
    Else
        strSql = " Select a.Id, a.��������, a.����, a.·������, a.�汾˵��, b.��������" & vbNewLine & _
                 " From ��׼·��Ŀ¼ A, ��׼·������ B" & vbNewLine & _
                 " Where a.Id = b.��׼·��id and Nvl(a.���,0)=[1] order by ��������,ID "
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngIndex)
    
    For i = 0 To rsTemp.RecordCount - 1
        '��ÿ�����ҿ�ʼ�ĵط���ӷ�����
        If strDept <> CStr(rsTemp!��������) Then
            Set objPreRecord = rptStPath.Records.Add()
            Set objItem = objPreRecord.AddItem(CStr(""))
            Set objItem = objPreRecord.AddItem(CStr(rsTemp!��������))
            Set objItem = objPreRecord.AddItem("")
            Set objItem = objPreRecord.AddItem("")
            Set objItem = objPreRecord.AddItem("")
            Set objItem = objPreRecord.AddItem("")
            objPreRecord.Tag = ""
            objPreRecord.Expanded = False
            '�����Ӽ�¼
            Set objRecord = objPreRecord.Childs.Add()
            Set objItem = objRecord.AddItem(CStr(rsTemp!ID))
            Set objItem = objRecord.AddItem("")
            Set objItem = objRecord.AddItem(CStr(rsTemp!���� & ""))
            Set objItem = objRecord.AddItem(CStr(rsTemp!·������ & ""))
            Set objItem = objRecord.AddItem(CStr(rsTemp!�汾˵�� & ""))
            Set objItem = objRecord.AddItem(CStr(rsTemp!�������� & ""))
            objRecord.Tag = CStr(rsTemp!ID)
            strDept = CStr(rsTemp!��������)
            
            If mlngStPathID = 0 Then
                mlngStPathID = rsTemp!ID
                objPreRecord.Expanded = True
            Else
                If rsTemp!ID = mlngStPathID Then
                    objPreRecord.Expanded = True
                End If
            End If

            rsTemp.MoveNext
        Else
            Set objRecord = objPreRecord.Childs.Add()
            Set objItem = objRecord.AddItem(CStr(rsTemp!ID))
            Set objItem = objRecord.AddItem("")
            Set objItem = objRecord.AddItem(CStr(rsTemp!���� & ""))
            Set objItem = objRecord.AddItem(CStr(rsTemp!·������ & ""))
            Set objItem = objRecord.AddItem(CStr(rsTemp!�汾˵�� & ""))
            Set objItem = objRecord.AddItem(CStr(rsTemp!�������� & ""))
            objRecord.Tag = CStr(rsTemp!ID)
            If rsTemp!ID = mlngStPathID Then
                objPreRecord.Expanded = True
            End If
            strDept = CStr(rsTemp!��������)
            rsTemp.MoveNext
        End If
    Next
    rptStPath.Populate
    '��λ��ѡ��ı�׼·��
    For i = 0 To rptStPath.Rows.count - 1
        If Val(rptStPath.Rows(i).Record.Tag) = mlngStPathID Then
            rptStPath.Rows(i).Selected = True
            Exit For
        End If
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadPathByID(ByVal lngId As Long, Optional ByVal blnReadData As Boolean, Optional ByVal lng��� As Long)
'���ܣ�����ѡ��ı�׼·��ID��ȡ���ݣ������ݱ���ż���·�����̣�·����������ͷ
'������lngID   ѡ���·��ID
'      blnReadData �Ƿ��ȡ��׼·����Ϣ���ڱ�׼·�����μ��ػ��߱�׼·���л�ʱ�Ǿ���Ҫ��ȡ��
'      lng���  0 ��׼·�����̣�1 ��1��2����2...
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, j As Long, k As Long
    Dim strSql As String, strFilter As String
    Dim strTilePos As String '��¼������λ�ø�ʽΪ������1��ʼλ�ã�����;����2��ʼλ�ã�����
    Dim lngColCount As Long, lng������ As Long, lngBeginRow As Long
    Dim lngRowCount As Long
    Dim strContent As String
    
    On Error GoTo errH
    
    If blnReadData Then
        'ɾ��ѡ������vs����
        vsPathTable.Delete
        For i = tbcStPath.ItemCount - 1 To 1 Step -1
            tbcStPath.RemoveItem (i)
        Next
        
        '���ر�׼����
        rtfPathCourse.Visible = False
        rtfPathCourse.Text = ""
        If mintMode = 1 Then
            strSql = "Select ����, ���� From ��׼����·������ Where ��׼·��id = [1] Order By ���"
        Else
            strSql = "Select ����, ���� From ��׼·������ Where ��׼·��id = [1] Order By ���"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngId)
        If rsTmp.RecordCount <> 0 Then
            For i = 1 To rsTmp.RecordCount
                strTilePos = strTilePos & ";" & Len(strContent) & "," & Len(rsTmp!����)
                strContent = strContent & rsTmp!���� & vbNewLine & vbNewLine & rsTmp!���� & vbNewLine & vbNewLine
                rsTmp.MoveNext
            Next
            rtfPathCourse.Text = strContent
        End If
               
        
        Call SetStPathCourceFont(Mid(strTilePos, 2)) '��������
        rtfPathCourse.Visible = True
        
        '��ȡ��������Ϣ
        If mintMode = 1 Then
            strSql = "Select a.����� �����, b.������, b.����ͷ, a.����, a.����" & vbNewLine & _
                    "From (Select �����, Max(�������) ����, Max(�׶����) ���� From ��׼����·���� Where ��׼·��id = [1] Group By �����) A, ��׼����·���� B" & vbNewLine & _
                    "Where b.��׼·��id =[1] And a.����� = b.����� And b.������� = 1 And b.�׶���� = 1" & vbNewLine & _
                    "Order By �����"
        Else
            strSql = "Select a.����� �����, b.������, b.����ͷ, a.����, a.����" & vbNewLine & _
                    "From (Select �����, Max(�������) ����, Max(�׶����) ���� From ��׼·���� Where ��׼·��id = [1] Group By �����) A, ��׼·���� B" & vbNewLine & _
                    "Where b.��׼·��id =[1] And a.����� = b.����� And b.������� = 1 And b.�׶���� = 1" & vbNewLine & _
                    "Order By �����"
        End If

        Set mrs��ͷ��Ϣ = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngId)
        
        '���ر�׼·����ѡ��
        If mrs��ͷ��Ϣ.RecordCount > 0 Then
            j = mrs��ͷ��Ϣ.RecordCount
            For i = 1 To j
                mrs��ͷ��Ϣ.Filter = "����� =" & i
                Call tbcStPath.InsertItem(i, mrs��ͷ��Ϣ!������, picPathTable.Hwnd, 0)
            Next
            '��ȡ������
            If mintMode = 1 Then
                strSql = "Select  �����, ������, ����ͷ, �������, ��������, �׶����, �׶�����, ·������" & vbNewLine & _
                    "From   ��׼����·����" & vbNewLine & _
                    "where ��׼·��id=[1]"
            Else
                strSql = "Select  �����, ������, ����ͷ, �������, ��������, �׶����, �׶�����, ·������" & vbNewLine & _
                    "From   ��׼·����" & vbNewLine & _
                    "where ��׼·��id=[1]"
            End If
            Set mrs�� = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngId)
        End If
    End If
    
    If lng��� <> 0 Then
        'û�б���Ϣ�����򲻼��ر���Ϣ
        mrs��.Filter = ""
        If mrs��.RecordCount = 0 Then tbcStPath.Item(0).Selected = True: Exit Sub
        mrs��ͷ��Ϣ.Filter = " ����� =" & lng���
        If mrs��ͷ��Ϣ.RecordCount = 0 Then tbcStPath.Item(0).Selected = True: Exit Sub
        
        '���ر���ͷ
        lbl��ͷ.Caption = ""
        lbl��ͷ.Caption = vbNewLine & mrs��ͷ��Ϣ!����ͷ
        
        With vsPathTable
            .Redraw = False
            .Rows = 0
            .Cols = 0
            'ȷ������
            lngColCount = Val(mrs��ͷ��Ϣ!���� & "") - 1
            'ȷ��������
            lng������ = Val(mrs��ͷ��Ϣ!���� & "")
            lngRowCount = IntEx(lngColCount / M_INT_STEPNUM) * lng������ + IntEx(lngColCount / M_INT_STEPNUM) - 1
            If lngRowCount = 1 And lngColCount = 1 Then
                .Rows = 0
                .Cols = 0
                Call SetVsStyle
                Call picPathTable_Resize '����lbl��ͷ��autoSize�������Ҫ����resize
                tbcStPath.Item(lng���).Selected = True
                Exit Sub
            Else
                .Rows = lngRowCount
                .Cols = IIf(lngColCount > M_INT_STEPNUM, M_INT_STEPNUM + 1, lngColCount + 1)
            End If
    
            For k = 1 To IntEx(lngColCount / M_INT_STEPNUM)
                lngBeginRow = (k - 1) * lng������ + (k - 1)
                For i = lngBeginRow To lngBeginRow + lng������ - 1
                    For j = 0 To .Cols - 1
                        'ÿ�����������ĵ�һ����Ԫ��Ϊʱ��
                        If i = lngBeginRow And j = 0 Then
                            .TextMatrix(i, j) = "ʱ��"
                        Else
                            If Not (i = lngBeginRow Or j = 0) Then
                                strFilter = "�����=" & lng��� & " and �������=" & i - lngBeginRow + 1 & " and �׶����=" & (k - 1) * 3 + j + 1
                                mrs��.Filter = strFilter
                                If mrs��.RecordCount = 1 Then
                                    .TextMatrix(i, j) = Nvl(mrs��!·������, " ")
                                    .TextMatrix(i, 0) = Replace(Replace(Replace(mrs��!�������� & "", Chr(13), ""), Chr(10), ""), " ", "")
                                    .TextMatrix(lngBeginRow, j) = mrs��!�׶����� & ""
                                End If
                            End If
                        End If
                    Next
                Next
            Next
            
            Call SetVsStyle
            .Redraw = True
            Call picPathTable_Resize '����lbl��ͷ��autoSize�������Ҫ����resize
        
            
        End With
    End If
    
    tbcStPath.Item(lng���).Selected = True
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetStPathCourceFont(ByVal strTilePos As String)
'���ܣ���RichTextBox������������
'���� strTilePos ��¼������λ�ø�ʽΪ������1��ʼλ�ã����ⳤ��;����2��ʼλ�ã����ⳤ��
    Dim arrtmp As Variant, i As Long
    
    On Error Resume Next
    
    If Len(Trim(strTilePos)) = 0 Then Exit Sub
    arrtmp = Split(Trim(strTilePos), ";")

    With rtfPathCourse

        For i = LBound(arrtmp) To UBound(arrtmp)
            .SelStart = Split(arrtmp(i), ",")(0)
            .SelLength = Split(arrtmp(i), ",")(1)
            .SelFontSize = 14
            .SelFontName = "����"
            .SelBold = True
            .SelLength = 0
        Next

        .SelStart = 0 '����ƶ�����ʼ
    End With
End Sub

Private Sub SetVsStyle()
'���ܣ������������ñ����ĵ�Ԫ��ĸ߶�����,�Լ�������ɫ�ȣ��Լ���Ԫ��ĺϲ���
    Dim i As Long, j As Long
    Dim lngmaxHeight As Long
    Dim strTmp As String

    With vsPathTable

        '�޸ķ������ƣ��׶Σ�����Ӵ־���
        .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, 0) = 4 '����
        .Cell(flexcpBackColor, 0, 0, .Rows - 1, 0) = &HE1FFE1

        .AutoResize = False
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .Cols - 1, False, 0) '�Զ�������С
        '���ý׶����壬��ɫ�����뷽ʽ
        For i = 0 To .Rows - 1
            If .TextMatrix(i, 0) = "ʱ��" Then
                .Cell(flexcpAlignment, i, 0, i, .Cols - 1) = 4
                .Cell(flexcpFontBold, i, 0, i, .Cols - 1) = False '���üӴ�ǰҪ������Ӵ�
                .Cell(flexcpFontBold, i, 0, i, .Cols - 1) = True
                .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HE1FFE1
            Else
                .Cell(flexcpAlignment, i, 1, i, .Cols - 1) = 0
            End If
        Next

        '��ȡͬһ����ߵĵ�Ԫ��߶ȸ�ֵ���и�
        For i = 0 To .Rows - 1
            If .TextMatrix(i, 0) <> "" Then
                For j = 0 To .Cols - 1
                    If j = 0 Then
                        lngmaxHeight = ComputerLines(.TextMatrix(i, j))
                    Else
                        lngmaxHeight = IIf(lngmaxHeight > ComputerLines(.TextMatrix(i, j)), lngmaxHeight, ComputerLines(.TextMatrix(i, j)))
                    End If
                Next
                .RowHeight(i) = lngmaxHeight * Me.TextHeight("��") * 1.5
            Else
                For j = 0 To .Cols - 1
                    .TextMatrix(i, j) = " " 'Ϊ�˺ϲ���Ԫ��
                Next
            End If
        Next
        '�ָ��е�Ԫ��ϲ����Լ��߿���ɫ����
        .MergeCells = flexMergeFree
        For i = 0 To .Rows - 1
            If .TextMatrix(i, 0) = " " Then
                Call .CellBorderRange(i, 0, i, .Cols - 1, &HFFFFFF, 1, 0, 1, 0, 1, 0)
                .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HFFFFFF
                .MergeRow(i) = True
            End If
        Next
        For i = 1 To .Cols - 1
            .ColWidth(i) = 4000
        Next
        .ColWidth(0) = 1500
        'ʵ�������϶��п�
        .FixedRows = 1
        Call .CellBorderRange(0, 0, 0, .Cols - 1, &H8000&, 0, 0, 1, 1, 1, 1)
    End With
End Sub

Private Function ComputerLines(ByVal strInput As String) As Long
'���ܣ����������ı��лس����ĸ���
'������  strInput   Ҫ����س������ַ���
'���أ�   �س����ĸ���

    Dim strTmp As String
    Dim count  As Long, lngPos As Long, lngLen As Long
    
    lngPos = InStr(strInput, Chr(13))
    lngLen = Len(strInput)
    strTmp = strInput
    
    Do While lngPos <> 0
        If Trim(strTmp) = "" Then Exit Do
        If lngPos + 1 <= lngLen Then
            strTmp = Mid(strTmp, lngPos + 1)
            count = count + 1
            lngPos = InStr(strTmp, Chr(13))
            lngLen = Len(strTmp)
        End If
    Loop
    
    ComputerLines = count + 2
End Function

Private Sub txtInput_GotFocus()
    Call zlControl.TxtSelAll(txtInput)
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call FuncFindSTPath
        txtInput.SetFocus
    End If
End Sub

Private Sub txtInput_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strTip As String
    
    strTip = "��·������\�����������" & vbCrLf & "����(Ctrl+F)" & vbCrLf & "������һ��(F3)"
    zlCommFun.ShowTipInfo txtInput.Hwnd, strTip, True
End Sub


