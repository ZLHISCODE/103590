VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPathManageOut 
   AutoRedraw      =   -1  'True
   Caption         =   "�����ٴ�·������"
   ClientHeight    =   7950
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   10890
   Icon            =   "frmPathManageOut.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7950
   ScaleWidth      =   10890
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   8385
      TabIndex        =   12
      Top             =   180
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog cdgFile 
      Left            =   3180
      Top             =   285
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   2250
      ScaleHeight     =   600
      ScaleWidth      =   660
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   210
      Visible         =   0   'False
      Width           =   660
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   1200
      Top             =   300
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathManageOut.frx":058A
            Key             =   "Path"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathManageOut.frx":0B24
            Key             =   "File"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathManageOut.frx":10BE
            Key             =   "branch"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathManageOut.frx":7920
            Key             =   "Merge"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6285
      Left            =   60
      ScaleHeight     =   6285
      ScaleWidth      =   3615
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1065
      Width           =   3615
      Begin XtremeReportControl.ReportControl rptPath 
         Height          =   3975
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   3375
         _Version        =   589884
         _ExtentX        =   5953
         _ExtentY        =   7011
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VSFlex8Ctl.VSFlexGrid vsFile 
         Height          =   810
         Left            =   225
         TabIndex        =   2
         Top             =   5265
         Width           =   3150
         _cx             =   5556
         _cy             =   1429
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
         BackColorSel    =   16571840
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   4
         GridLinesFixed  =   5
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   285
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPathManageOut.frx":E182
         ScrollTrack     =   -1  'True
         ScrollBars      =   0
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
   Begin VB.Frame fraLR 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5400
      Left            =   3720
      MousePointer    =   9  'Size W E
      TabIndex        =   6
      Top             =   1485
      Width           =   45
   End
   Begin XtremeSuiteControls.TabControl tbcContent 
      Height          =   4155
      Left            =   3930
      TabIndex        =   4
      Top             =   3225
      Width           =   6735
      _Version        =   589884
      _ExtentX        =   11880
      _ExtentY        =   7329
      _StockProps     =   64
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2385
      Left            =   3930
      ScaleHeight     =   2385
      ScaleWidth      =   6720
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   840
      Width           =   6720
      Begin VSFlex8Ctl.VSFlexGrid vsgIllness 
         Height          =   855
         Left            =   480
         TabIndex        =   13
         Top             =   1440
         Width           =   5535
         _cx             =   9763
         _cy             =   1508
         Appearance      =   0
         BorderStyle     =   0
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
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
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
         FillStyle       =   1
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
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����ÿ��ң�"
         Height          =   180
         Index           =   0
         Left            =   165
         TabIndex        =   10
         Top             =   555
         Width           =   1080
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������������������������������������������������������������"
         Height          =   180
         Index           =   1
         Left            =   330
         MouseIcon       =   "frmPathManageOut.frx":E1BF
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   780
         Width           =   5475
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���Ӧ���֣�"
         Height          =   180
         Index           =   0
         Left            =   165
         TabIndex        =   8
         Top             =   1215
         Width           =   1080
      End
      Begin VB.Label lbl˵�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "˵������������������������������������������������������������������"
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   6210
         WordWrap        =   -1  'True
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   7590
      Width           =   10890
      _ExtentX        =   19209
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPathManageOut.frx":E311
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16298
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   270
      Top             =   165
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPathManageOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mfrmDesign   As frmPathDesignOut
Attribute mfrmDesign.VB_VarHelpID = -1
Private WithEvents mfrmContent  As frmPathDesignOut
Attribute mfrmContent.VB_VarHelpID = -1
Private WithEvents mfrmEdit     As frmPathEditOut       '�������޸�·���Ĵ���
Attribute mfrmEdit.VB_VarHelpID = -1
Private mstr��������            As String               '��ǰѡ����ٴ�·���Ķ�Ӧ��������
Private mstrPrivs               As String
Private mstrDictPrivs           As String
Private mlngModul               As Long
Private zlAppTool               As Object

Private Enum COL_LIST
    COL_ID = 0
    COL_ͼ�� = 1
    COL_��֧ = 2
    COL_�к� = 3
    COL_���� = 4
    COL_���� = 5
    COL_���� = 6
    COL_�����Ա� = 7
    COL_�������� = 8
    COL_˵�� = 9
    COL_ͨ�� = 10
    COL_���°汾 = 11
    COL_ƴ������ = 12
End Enum

Private Sub FuncPathNew()
'����: ���������ٴ�·��
    Dim str���� As String

    If InStr(mstrPrivs, "������·������") > 0 Then
        
    ElseIf InStr(mstrPrivs, "30������·��") > 0 Then
        If rptPath.Records.count >= 30 Then
            MsgBox "�Ѵﵽ�����Ȩ�����·����������������������", vbInformation, gstrSysName
            Exit Sub
        End If
    ElseIf InStr(mstrPrivs, "5������·��") > 0 Then
        If rptPath.Records.count >= 5 Then
            MsgBox "�Ѵﵽ�����Ȩ�����·����������������������", vbInformation, gstrSysName
            Exit Sub
        End If
    Else
        MsgBox "������ȷ��Ȩ�����·������������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
        Exit Sub
    End If

    If rptPath.SelectedRows.count > 0 Then
        If rptPath.SelectedRows(0).GroupRow Then
            str���� = rptPath.SelectedRows(0).Childs(0).Record(COL_����).Value
        Else
            str���� = rptPath.SelectedRows(0).Record(COL_����).Value
        End If
    End If
    mfrmEdit.ShowEdit Me, mstrPrivs, , str����                                      '�¼�����ˢ��
End Sub

Private Sub FuncPathModify()
'����: �޸������ٴ�·��
    mfrmEdit.ShowEdit Me, mstrPrivs, rptPath.SelectedRows(0).Record(COL_ID).Value   '�¼�����ˢ��
End Sub

Private Sub FuncPathDelete()
'����: ɾ�������ٴ�·��
    Dim strSql As String

    With rptPath.SelectedRows(0)
        If MsgBox("ȷʵҪɾ���ٴ�·��""" & .Record(COL_����).Value & """��", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName) = vbNo Then Exit Sub

        strSql = "Zl_����·��Ŀ¼_Delete(" & .Record(COL_ID).Value & ")"
        On Error GoTo errH
        zlDatabase.ExecuteProcedure strSql, Me.Caption
        
        Call RefreshData
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncPathFileDelete()
'����: ɾ�������ٴ�·���ļ�
    Dim strSql As String

    With vsFile
        If MsgBox("ȷʵҪɾ���ļ�""" & .TextMatrix(.Row, 1) & """��", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName) = vbNo Then Exit Sub

        strSql = "Zl_����·���ļ�_Delete(" & rptPath.SelectedRows(0).Record(COL_ID).Value & ",'" & .TextMatrix(.Row, 1) & "')"
        On Error GoTo errH
        zlDatabase.ExecuteProcedure strSql, Me.Caption
        On Error GoTo 0

        .RemoveItem .Row
        .Height = .Height - .RowHeightMin
        Call picList_Resize
        Call vsFile_AfterRowColChange(0, 0, .Row, .Col)
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncPathFileView()
'���ܣ����ٴ�·���ļ��鿴
    Dim strFile As String
    Dim lngRetu As Long, strInfo As String

    Screen.MousePointer = 11
    
    On Error GoTo errH
    strFile = gobjFile.GetSpecialFolder(TemporaryFolder) & "\" & vsFile.TextMatrix(vsFile.Row, 1)
    If gobjFile.FileExists(strFile) Then gobjFile.DeleteFile strFile, True

    strFile = Sys.ReadLob(glngSys, 26, rptPath.SelectedRows(0).Record(COL_ID).Value & "," & vsFile.TextMatrix(vsFile.Row, 1), strFile)
    If Not gobjFile.FileExists(strFile) Then
        MsgBox "�ļ����ݶ�ȡʧ�ܣ�", vbInformation, gstrSysName:
        Screen.MousePointer = 0: Exit Sub
    End If

    lngRetu = ShellExecute(Me.Hwnd, "open", strFile, "", "", SW_SHOWNORMAL)
    If lngRetu <= 32 Then
        Select Case lngRetu
            Case 2
                strInfo = "����Ĺ���"
            Case 29
                strInfo = "����ʧ��"
            Case 30
                strInfo = "����Ӧ�ó�ʽæµ��..."
            Case 31
                strInfo = "û�й����κ�Ӧ�ó�ʽ"
            Case Else
                strInfo = "�޷�ʶ��Ĵ���"
        End Select
        MsgBox "�ļ���ʱ����" & vbCrLf & vbCrLf & vbTab & strInfo, vbExclamation, gstrSysName
    End If

    Screen.MousePointer = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncPathFileNew(ByVal lngModle As Long)
'������lngModle=0 ������ͨ�ļ���=1 �������߰�·����
    Dim arrSQL() As String
    Dim strFile As String
    Dim strFileName As String
    Dim i As Long
    Dim blnTrans As Boolean

    cdgFile.DialogTitle = "ѡ��Ҫ��ӵ������ٴ�·���ļ�"
    If lngModle = 0 Then
        cdgFile.Filter = "�����ļ�|*.*"
    Else
        cdgFile.Filter = "Word�ĵ�(*.doc;*.docx)|*.doc;*.docx"
        For i = 1 To vsFile.Rows - 1
            If vsFile.Cell(flexcpForeColor, i, 1) = &HFF0000 Then
                MsgBox "��ǰ·���Ѿ����ڻ��߰�·�����ļ�����ɾ�����ٽ�����ӡ�", vbInformation, Me.Caption
                Exit Sub
            End If
        Next
    End If
    cdgFile.Flags = &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
    cdgFile.InitDir = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "����·���ļ�Ŀ¼")
    cdgFile.CancelError = True
    On Error Resume Next
    cdgFile.ShowOpen
    If Err.Number <> 0 Then
        Err.Clear: Exit Sub
    End If
    On Error GoTo 0
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "����·���ļ�Ŀ¼", gobjFile.GetFile(cdgFile.FileName).ParentFolder.Path
    strFile = cdgFile.FileName                              '����·��
    strFileName = gobjFile.GetFile(cdgFile.FileName).Name

    '����ļ���С������3M
    If gobjFile.GetFile(strFile).Size / 1024 / 1024 > 3 Then
        MsgBox "�ļ��ߴ�̫��(����3M)������ļ������ʵ������������ӡ�", vbInformation, gstrSysName
        Exit Sub
    End If

    Screen.MousePointer = 11

    ReDim arrSQL(0)
    arrSQL(0) = "Zl_����·���ļ�_Insert(" & rptPath.SelectedRows(0).Record(COL_ID).Value & ",'" & strFileName & "'," & lngModle & ")"
    If Not Sys.GetLobSql(glngSys, 26, rptPath.SelectedRows(0).Record(COL_ID).Value & "," & strFileName, strFile, arrSQL()) Then
        MsgBox "�ļ����ʧ�ܣ�", vbExclamation, gstrSysName
        Screen.MousePointer = 0
        Exit Sub
    End If

    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    For i = LBound(arrSQL) To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(arrSQL(i), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False
    On Error GoTo 0

    Call LoadPathFile(rptPath.SelectedRows(0).Record(COL_ID).Value) 'ˢ��

    Screen.MousePointer = 0
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncPathTableOutputAll()
'���ܣ����ȫ���ٴ�·����Excel
    Dim lngCount As Long, i As Long
    Dim objRow As ReportRow
    Dim objControl As CommandBarControl

    lngCount = rptPath.Records.count
    If lngCount = 0 Then
        MsgBox "��ǰû�п��������·����", vbInformation, gstrSysName
    Else
        If MsgBox("����" & lngCount & "��·������ȷ��Ҫȫ�������Excel��", vbQuestion + vbOKCancel + vbDefaultButton1, gstrSysName) = vbCancel Then
            Exit Sub
        End If
    End If

    For i = 0 To rptPath.Rows.count - 1
        If Not rptPath.Rows(i).GroupRow Then
            Set objRow = rptPath.Rows(i)
            Set rptPath.FocusedRow = objRow                 '����ѡ������ʾ�ڿɼ�����,������SelectionChanged�¼�

            Set objControl = cbsMain.FindControl(, conMenu_File_Excel, True, True)
            If Not objControl Is Nothing Then
                Call mfrmContent.zlExecuteCommandBars(objControl, True)
            End If
        End If
    Next
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim objRow As ReportRow, i As Long
    Dim lng·��ID As Long
    Dim blnTmp As Boolean
    Dim str���� As String
    Dim str���� As String
    Dim frmSub As Form

    If Control.ID <> 0 And Control.ID <> conMenu_View_FindNext Then
        If cbsMain.FindControl(, Control.ID, True, True) Is Nothing Then Exit Sub
    End If

    Select Case Control.ID
        Case conMenu_File_ExportToXML * 10# + 1    '��������XML
            Call FuncExportToXMLBatch
        Case conMenu_File_ExportToXML * 10# + 2    '��������XML
            Call FuncImportFromXMLBatch
        Case conMenu_File_ExportToXML * 10# + 3    '�����׼·��
            Set frmSub = New frmStandardPathRef
            If frmSub.ShowMe(gfrmMain, 0, 1, 1) Then
                'ˢ��
                Call RefreshData
            End If
        Case conMenu_File_BatPrint              'ȫ�������Excel
            Call FuncPathTableOutputAll
        Case conMenu_Edit_NewItem               '����
            Call FuncPathNew
        Case conMenu_Edit_Modify                '�޸�
            Call FuncPathModify
        Case conMenu_Edit_Delete                'ɾ��
            Call FuncPathDelete
        Case conMenu_Edit_Archive * 10# + 1     '�����ļ�
            Call FuncPathFileNew(0)
        Case conMenu_Edit_Archive * 10# + 2     '�����ļ�
            Call FuncPathFileNew(1)
        Case conMenu_Edit_Archive * 10# + 3     '�鿴�ļ�
            Call FuncPathFileView
        Case conMenu_Edit_Archive * 10# + 4     'ɾ���ļ�
            Call FuncPathFileDelete
        Case conMenu_Tool_Define                'ͼ������
            If frmIconManage.ShowMe(Me) Then
                Call rptPath_SelectionChanged
            End If
        Case conMenu_Tool_Option                '�ֵ����
            If zlAppTool Is Nothing Then Set zlAppTool = CreateObject("zl9AppTool.clsAppTool")
            Call zlAppTool.zlAppointDict("·���������,·���������,������쳣��ԭ��", glngSys)
        Case conMenu_Edit_Report                '�����ǼǱ�
            Call frmPathOutDefinition.ShowMe(Me, rptPath.SelectedRows(0).Record(COL_ID).Value, rptPath.SelectedRows(0).Record(COL_����).Value, 1)
        Case conMenu_Edit_Compend               '���
            If InStr(mstrPrivs, "������·������") > 0 Then
                'Do Nothing
            ElseIf InStr(mstrPrivs, "30������·��") > 0 Then
                If rptPath.Records.count > 30 Then
                    MsgBox "�Ѵﵽ�����Ȩ�����·������������ɾ��������ٴ�·����", vbInformation, gstrSysName
                    Exit Sub
                End If
            ElseIf InStr(mstrPrivs, "5������·��") > 0 Then
                If rptPath.Records.count > 5 Then
                    MsgBox "�Ѵﵽ�����Ȩ�����·������������ɾ��������ٴ�·����", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
            Call mfrmDesign.ShowDesign(Me, rptPath.SelectedRows(0).Record(COL_ID).Value, mstrPrivs, mstr��������)
            Call RefreshData    'Ϊ����ƽ��������汾��ˢ��·��Ŀ¼��δ���·������ɫ
'        Case conMenu_Edit_Adjust  ' �����Ľ�
'            If rptPath.SelectedRows(0).Record(COL_���°汾).Value = 0 Then
'                MsgBox "��·��Ϊδ������ù����½�·��,����ִ�и����Ľ����ܡ�", vbInformation, gstrSysName
'                Exit Sub
'            End If
'            str���� = rptPath.SelectedRows(0).Record(COL_����).Value
'            Call frmPathImprove.ShowMe(Me, rptPath.SelectedRows(0).Record(COL_ID).Value, str����, str����, blnTmp)
'            If blnTmp Then Call RefreshData(str����, str����)  'ˢ������
'        Case conMenu_Edit_BatExecute                            '��������
'            Call frmPathItemBatReplace.ShowMe(Me, mstrPrivs)
        Case conMenu_View_Find                                  '����
            If Me.ActiveControl Is txtFind Then
                txtFind.SetFocus                                '��ʱ��Ҫ��λһ��
                If txtFind.Text <> "" Then
                    Call FuncFindPath
                End If
            Else
                txtFind.SetFocus
            End If
        Case conMenu_View_FindNext          '������һ��
            If txtFind.Text = "" Then
                txtFind.SetFocus
            Else
                Call FuncFindPath(True)
            End If
        Case conMenu_View_ToolBar_Button    '������
            For i = 2 To cbsMain.count
                Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
            Next
            Me.cbsMain.RecalcLayout
        Case conMenu_View_ToolBar_Text    '��ť����
            For i = 2 To cbsMain.count
                For Each objControl In Me.cbsMain(i).Controls
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                Next
            Next
            Me.cbsMain.RecalcLayout
        Case conMenu_View_ToolBar_Size    '��ͼ��
            Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
            Me.cbsMain.RecalcLayout
        Case conMenu_View_StatusBar    '״̬��
            Me.stbThis.Visible = Not Me.stbThis.Visible
            Me.cbsMain.RecalcLayout
        Case conMenu_View_StPath    '�鿴��׼·���ο�
            Call frmStPathList.ShowMe(Me, mstr��������, 1)
        Case conMenu_View_Expend_CurCollapse    '�۵���ǰ��
            If rptPath.SelectedRows.count > 0 Then
                If rptPath.SelectedRows(0).GroupRow Then
                    rptPath.SelectedRows(0).Expanded = False
                ElseIf Not rptPath.SelectedRows(0).ParentRow Is Nothing Then
                    If rptPath.SelectedRows(0).ParentRow.GroupRow Then
                        rptPath.SelectedRows(0).ParentRow.Expanded = False
                    End If
                End If
            End If
            '���۵���λ��������,�����Զ�������¼�
            Call rptPath_SelectionChanged
        Case conMenu_View_Expend_CurExpend    'չ����ǰ��
            If rptPath.SelectedRows.count > 0 Then
                rptPath.SelectedRows(0).Expanded = True
            End If
        Case conMenu_View_Expend_AllCollapse    '�۵�������
            For Each objRow In rptPath.Rows
                If objRow.GroupRow Then objRow.Expanded = False
            Next
            '���۵���λ��������,�����Զ�������¼�
            Call rptPath_SelectionChanged
        Case conMenu_View_Expend_AllExpend    'չ��������
            For Each objRow In rptPath.Rows
                If objRow.GroupRow Then objRow.Expanded = True
            Next
        Case conMenu_View_Refresh           'ˢ��
            Call RefreshData
        Case conMenu_Help_Web_Home          'Web�ϵ�����
            Call zlHomePage(Me.Hwnd)
        Case conMenu_Help_Web_Forum         '������̳
            Call zlWebForum(Me.Hwnd)
        Case conMenu_Help_Web_Mail          '���ͷ���
            Call zlMailTo(Me.Hwnd)
        Case conMenu_Help_About             '����
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case conMenu_Help_Help              '����
            Call ShowHelp(App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_File_Exit              '�˳�
            Unload Me
        Case Else
            If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
                If rptPath.SelectedRows.count > 0 Then
                    If Not rptPath.SelectedRows(0).GroupRow Then
                        lng·��ID = rptPath.SelectedRows(0).Record(COL_ID).Value
                    End If
                End If
                'ִ�з�������ǰģ��ı���
                Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, "·��ID=" & lng·��ID)
            Else
                Call mfrmContent.zlExecuteCommandBars(Control)
                Select Case Control.ID
                    Case conMenu_Edit_Audit, conMenu_Edit_Untread    '���,ȡ�����
                        Call RefreshData
                    Case conMenu_Edit_Stop, conMenu_Edit_Reuse      'ͣ��,ȡ��ͣ��
                        Call RefreshData
                End Select
            End If
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    On Error Resume Next

    With Me.picList
        .Left = lngLeft: .Top = lngTop
        .Height = lngBottom - lngTop
    End With

    With Me.fraLR
        .Left = Me.picList.Left + Me.picList.Width
        .Top = Me.picList.Top
        .Height = Me.picList.Height
    End With

    With Me.PicInfo
        .Left = fraLR.Left + fraLR.Width
        .Top = fraLR.Top
        .Width = lngRight - .Left
    End With
    Call ResizeInfoPane

    With Me.tbcContent
        .Left = PicInfo.Left
        .Top = PicInfo.Top + PicInfo.Height
        .Width = PicInfo.Width
        .Height = lngBottom - .Top
    End With

    Me.Refresh
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean
    Dim lng·��ID As Long

    If rptPath.SelectedRows.count > 0 Then
        If Not rptPath.SelectedRows(0).GroupRow Then
            lng·��ID = rptPath.SelectedRows(0).Record(COL_ID).Value
        End If
    End If

    Select Case Control.ID
        '��������XML���Ӵ������ж�
        Case conMenu_File_ExportToXML * 10# + 1    '��������XML
            If InStr(mstrPrivs, "����XML") = 0 Then
                Control.Visible = False
            Else
                Control.Enabled = rptPath.Records.count > 0
            End If
        Case conMenu_File_ExportToXML * 10# + 2    '��������XML
            If InStr(mstrPrivs, "����XML") = 0 Then
                Control.Visible = False
            End If
        Case conMenu_Edit_NewItem    '����
            If InStr(mstrPrivs, "��ɾ��") = 0 Then
                Control.Visible = False
            End If
        Case conMenu_Edit_Modify    '�޸�
            If InStr(mstrPrivs, "��ɾ��") = 0 Then
                Control.Visible = False
            Else
                Control.Enabled = lng·��ID <> 0
            End If
        Case conMenu_Edit_Delete    'ɾ��
            If InStr(mstrPrivs, "��ɾ��") = 0 Then
                Control.Visible = False
            Else
                Control.Enabled = lng·��ID <> 0
            End If
        Case conMenu_Edit_Archive * 10# + 1, conMenu_Edit_Archive * 10# + 2    '�����ļ�,�����ٴ�·����
            If InStr(mstrPrivs, "��ɾ��") = 0 Then
                Control.Visible = False
            Else
                Control.Enabled = lng·��ID <> 0
            End If
        Case conMenu_Edit_Archive * 10# + 3    '�鿴�ļ�
            Control.Enabled = lng·��ID <> 0 And vsFile.Rows > vsFile.FixedRows And vsFile.Row >= vsFile.FixedRows
        Case conMenu_Edit_Archive * 10# + 4    'ɾ���ļ�
            Control.Enabled = lng·��ID <> 0 And vsFile.Rows > vsFile.FixedRows And vsFile.Row >= vsFile.FixedRows
        Case conMenu_Edit_Report  '�����ǼǱ�
            If InStr(mstrPrivs, "�����ǼǱ����") = 0 Then
                Control.Visible = False
            End If
            blnEnabled = False
            If rptPath.SelectedRows.count > 0 Then
                If Not rptPath.SelectedRows(0).GroupRow Then
                   blnEnabled = True
                End If
            End If
            Control.Enabled = blnEnabled
'        Case conMenu_Edit_Adjust, conMenu_Edit_BatExecute '�����Ľ�,'��������
'            If Control.ID = conMenu_Edit_Adjust Then
'                If InStr(mstrPrivs, "�ٴ�·�������Ľ�") = 0 Then
'                    Control.Visible = False
'                End If
'            ElseIf Control.ID = conMenu_Edit_BatExecute Then
'                If InStr(mstrPrivs, "·�������") = 0 Then
'                    Control.Visible = False
'                End If
'            End If
'            If Control.Visible Then
'                If rptPath.SelectedRows.count > 0 Then
'                    If Not rptPath.SelectedRows(0).GroupRow Then
'                        Control.Enabled = True
'                    Else
'                        Control.Enabled = False
'                    End If
'                End If
'            End If
        Case conMenu_Tool_Define    'ͼ������
            If InStr(mstrPrivs, "ͼ������") = 0 Then
                Control.Visible = False
            End If
        Case conMenu_Tool_Option    '�ֵ����
            If InStr(mstrDictPrivs, "����") = 0 Then
                Control.Visible = False
            End If
        Case conMenu_View_ToolBar_Button    '������
            If cbsMain.count >= 2 Then
                Control.Checked = Me.cbsMain(2).Visible
            End If
        Case conMenu_View_ToolBar_Text    'ͼ������
            If cbsMain.count >= 2 Then
                Control.Checked = Not (Me.cbsMain(2).Controls(1).Style = xtpButtonIcon)
            End If
        Case conMenu_View_ToolBar_Size    '��ͼ��
            Control.Checked = Me.cbsMain.Options.LargeIcons
        Case conMenu_View_FindNext    '������һ��
            Control.Visible = False
        Case conMenu_View_StatusBar    '״̬��
            Control.Checked = Me.stbThis.Visible
        Case conMenu_View_Expend_CurExpend    'չ����ǰ��
            blnEnabled = False
            If rptPath.SelectedRows.count > 0 Then
                If rptPath.SelectedRows(0).GroupRow Then
                    blnEnabled = Not rptPath.SelectedRows(0).Expanded
                End If
            End If
            Control.Enabled = blnEnabled
        Case conMenu_View_Expend_CurCollapse    '�۵���ǰ��
            blnEnabled = False
            If rptPath.SelectedRows.count > 0 Then
                If rptPath.SelectedRows(0).GroupRow Then
                    blnEnabled = rptPath.SelectedRows(0).Expanded
                ElseIf Not rptPath.SelectedRows(0).ParentRow Is Nothing Then
                    If rptPath.SelectedRows(0).ParentRow.GroupRow Then
                        blnEnabled = rptPath.SelectedRows(0).ParentRow.Expanded
                    End If
                End If
            End If
            Control.Enabled = blnEnabled
        Case conMenu_View_Expend    '�۵�/չ����
            Control.Enabled = rptPath.GroupsOrder.count > 0 And rptPath.Rows.count > 0
        Case Else
            Call mfrmContent.zlUpdateCommandBars(Control)
    End Select
End Sub

Private Sub Form_Load()
    mstrPrivs = gstrPrivs
    mlngModul = glngModul
    '��ȡ�ֵ����Ȩ�޴�
    mstrDictPrivs = GetPrivFunc(0, 11)
    Call zlCommFun.SetWindowsInTaskBar(Me.Hwnd, False)

    Set mfrmEdit = New frmPathEditOut
    Set mfrmDesign = New frmPathDesignOut
    Set mfrmContent = New frmPathDesignOut
    'ReportControl
    '-----------------------------------------------------
    Call InitReportColumn
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True    '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    Call MainDefCommandBar

    'TabControl
    '-----------------------------------------------------
    With Me.tbcContent
        With .PaintManager
            .Appearance = xtpTabAppearanceVisio
            .Color = xtpTabColorOffice2003
        End With
        .InsertItem 0, "�ٴ�·����", mfrmContent.Hwnd, 0
    End With

    '�������
    '-----------------------------------------------------
    With vsFile
        .RowHeight(0) = 315
        Set .Cell(flexcpPicture, 0, 0) = img16.ListImages("File").Picture
        .Cell(flexcpPictureAlignment, 0, 0) = 7
        .TextMatrix(0, 1) = "·���ļ�����,��ɫ��ʾ·����(���߰�)"
        .Cell(flexcpFontBold, 0, 1) = True
    End With

    '��Ӧ����
    '---------------------------------------------------------
    Call InitVsgIllness                         '��ʼ����Ӧ���ֵ�VSF�ؼ�
    '
    Call RestoreWinState(Me, App.ProductName)

    Call RefreshData                            '���ݵ�ǰ���õ�������ȡ�ٴ�·��Ŀ¼����
End Sub

Private Sub MainDefCommandBar()
'���ܣ������ڲ˵����岿��
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    Dim lngCount As Long

    '�˵�����
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        Set objControl = .Add(xtpControlButton, conMenu_File_BatPrint, "ȫ�������&Excel��")
        Set objControl = .Add(xtpControlButton, conMenu_File_ExportToXML, "����&XML�ļ���")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_ExportToXML * 10# + 1, "��������XML�ļ���")
        Set objControl = .Add(xtpControlButton, conMenu_File_ExportToXML * 10# + 2, "��������XML�ļ���")
        'Set objControl = .Add(xtpControlButton, conMenu_File_ExportToXML * 10# + 3, "�����׼·��")
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)")
        objControl.BeginGroup = True
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����(&N)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�(&M)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&D)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Archive, "�ļ�(&F)")
        objPopup.BeginGroup = True
        objPopup.IconId = conMenu_Manage_Report
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Archive * 10# + 1, "�����ͨ�ļ�(&1)")
            objControl.IconId = conMenu_Edit_NewItem
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Archive * 10# + 2, "���·����(���߰�)(&2)")
            objControl.IconId = 3205
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Archive * 10# + 3, "�鿴�ļ�(&3)")
            objControl.IconId = conMenu_Tool_Search
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Archive * 10# + 4, "ɾ���ļ�(&4)")
            objControl.IconId = conMenu_Edit_Delete
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Compend, "���(&S)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Report, "�����ǼǱ����(&P)")
'        Set objControl = .Add(xtpControlButton, conMenu_Edit_Adjust, "�����Ľ�(&U)")
'        objControl.BeginGroup = True
'        Set objControl = .Add(xtpControlButton, conMenu_Edit_BatExecute, "��������(&B)")
'        objControl.IconId = conMenu_Apply_AllCard
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Audit, "���(&A)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "ȡ�����(&R)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Stop, "ͣ��(&S)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "ȡ��ͣ��(&Z)")
        objControl.IconId = conMenu_Edit_Untread
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "����(&T)", -1, False)
    objMenu.ID = conMenu_ToolPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Define, "ͼ������(&I)")
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Option, "�ֵ����(&D)")
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StPath, "��׼·���ο�")
        objControl.BeginGroup = True
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_Expend, "չ��/�۵���(&X)"):
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_AllCollapse, "�۵�������(&L)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_AllExpend, "չ��������(&X)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_CurCollapse, "�۵���ǰ��(&C)", -1, False)
            objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_CurExpend, "չ����ǰ��(&E)", -1, False)
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_FindNext, "������һ��(&C)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)")
        objControl.BeginGroup = True
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "��̳(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��")
        objControl.BeginGroup = True
    End With

    '���˵��Ҳ�Ĳ���
    With cbsMain.ActiveMenuBar.Controls
        Set objControl = .Add(xtpControlLabel, 0, "����")
        objControl.IconId = conMenu_View_Find
        objControl.Flags = xtpFlagRightAlign
        Set objCustom = .Add(xtpControlCustom, conMenu_View_Find, "")
        objCustom.Handle = txtFind.Hwnd
        objCustom.Flags = xtpFlagRightAlign
    End With

    '����������:������������
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("������", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Compend, "���")
'        Set objControl = .Add(xtpControlButton, conMenu_Edit_Adjust, "�����Ľ�")
'        objControl.BeginGroup = True
'        Set objControl = .Add(xtpControlButton, conMenu_Edit_BatExecute, "��������")
'        objControl.IconId = conMenu_Apply_AllCard
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Audit, "���"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "ȡ�����")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Stop, "ͣ��"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "ȡ��ͣ��")
        objControl.IconId = conMenu_Edit_Untread
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "����")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With

    '����һЩ�������ȼ���
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem                     '����
        .Add FCONTROL, vbKeyM, conMenu_Edit_Modify                      '�޸�
        .Add FSHIFT, vbKeyDelete, conMenu_Edit_Delete                   'ɾ��
        .Add FCONTROL, vbKeyD, conMenu_Edit_Compend                     '���/�޶�
        .Add FCONTROL, vbKeyU, conMenu_Edit_Audit                       '���
        .Add FCONTROL, vbKeyR, conMenu_Edit_Stop                        'ͣ��
        .Add FCONTROL, vbKeyF, conMenu_View_Find                        '����
        .Add 0, vbKeyF3, conMenu_View_FindNext                          '������һ��
        .Add FCONTROL, vbKeyAdd, conMenu_View_Expend_AllExpend          'չ��������
        .Add FCONTROL, vbKeySubtract, conMenu_View_Expend_AllCollapse   '�۵�������
        .Add FCONTROL, vbKeyP, conMenu_File_Print                       '��ӡ
        .Add 0, vbKeyF5, conMenu_View_Refresh                           'ˢ��
        .Add 0, vbKeyF1, conMenu_Help_Help                              '����
    End With

    '����һЩ�����Ĳ���������
    '-----------------------------------------------------
    With cbsMain.Options
        .AddHiddenCommand conMenu_File_PrintSet         '��ӡ����
        .AddHiddenCommand conMenu_File_Excel            '�����Excel
    End With

    '�ָ����̶���һЩ�˵�����
    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.SetIconSize 16, 16
    For lngCount = 2 To cbsMain.count
        cbsMain(lngCount).ContextMenuPresent = False
        cbsMain(lngCount).ShowTextBelowIcons = False
        cbsMain(lngCount).EnableDocking xtpFlagHideWrap
        For Each objControl In cbsMain(lngCount).Controls
            objControl.Style = xtpButtonIconAndCaption
        Next
    Next

    '��ȡ��������ģ��ı���(��������ģ���)
    '-----------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs)
End Sub

Private Sub InitReportColumn()
'��ʼ��ReportControl�ؼ�
    Dim objCol As ReportColumn

    With rptPath
        '����˳�������(�������Ϊ����)�ı��,Ҫ��Find(�к�)��ItemIndex������,���Կ���Record(�к�)����������
        Set objCol = .Columns.Add(COL_ID, "", 0, False)
        objCol.Visible = False
        Set objCol = .Columns.Add(COL_ͼ��, "", 18, False)
        objCol.Sortable = False
        objCol.AllowDrag = False
        objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Columns.Add(COL_��֧, "", 18, False)
        objCol.Sortable = False
        objCol.AllowDrag = False
        objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Columns.Add(COL_�к�, "�к�", 35, True)
        objCol.AllowDrag = False
        objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Columns.Add(COL_����, "����", 80, True)
        objCol.Visible = False
        Set objCol = .Columns.Add(COL_����, "����", 35, True)
        objCol.Groupable = False
        Set objCol = .Columns.Add(COL_����, "����", 150, True)
        objCol.Groupable = False
        Set objCol = .Columns.Add(COL_�����Ա�, "�����Ա�", 55, True)
        objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Columns.Add(COL_��������, "��������", 55, True)
        Set objCol = .Columns.Add(COL_˵��, "", 0, False)
        objCol.Visible = False
        Set objCol = .Columns.Add(COL_ͨ��, "", 0, False)
        objCol.Visible = False
        Set objCol = .Columns.Add(COL_���°汾, "", 0, False)
        objCol.Visible = False
        Set objCol = .Columns.Add(COL_ƴ������, "", 0, False)
        objCol.Visible = False

        For Each objCol In .Columns
            objCol.Editable = False
        Next

        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ���ٴ�·��..."
        End With
        .AutoColumnSizing = False
        .AllowColumnRemove = False
        .ShowGroupBox = True
        .ShowItemsInGroups = False
        .PreviewMode = True
        .MultipleSelection = False              '������SelectionChanged�¼�
        .SetImageList Me.img16

        .GroupsOrder.Add .Columns(COL_����)
        .GroupsOrder(0).SortAscending = True    '����֮��,��������в���ʾ,�����е������ǲ����

        '����֮�����ʧȥ��¼���е�˳��,���ǿ�м���������
        .SortOrder.Add .Columns(COL_����)
        .SortOrder(0).SortAscending = True
        .SortOrder.Add .Columns(COL_����)
        .SortOrder(1).SortAscending = True
    End With
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)

    Unload mfrmDesign
    Set mfrmDesign = Nothing
    
    Unload mfrmContent
    Set mfrmContent = Nothing

    Unload mfrmEdit
    Set mfrmEdit = Nothing
End Sub

Private Sub fraLR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next

    If Button = 1 Then
        If picList.Width + X < 2000 Or PicInfo.Width - X < 3000 Then Exit Sub

        fraLR.Left = fraLR.Left + X
        picList.Width = picList.Width + X

        Call Form_Resize
    End If
End Sub

Private Sub lbl����_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'���ң��������ȥ֮������»��ߣ���ɫ����
    If Index = 1 And lbl����(Index).Caption <> "" Then
        Me.lbl����(1).Font.Underline = True
        Me.lbl����(1).ForeColor = RGB(0, 0, 128)
    End If
End Sub

Private Sub mfrmDesign_DataChanged(ByVal ·��ID As Long)
'ˢ��·������Ϣ
    Call mfrmContent.zlRefresh(·��ID, mstrPrivs, lbl����(1).Caption, vsgIllness.Tag)
End Sub

Private Sub mfrmEdit_AfterSave(ByVal ���� As String, ByVal ���� As String)
    Call RefreshData(����, ����)
End Sub

Private Sub picInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'�������ȥ���ȡ���»��ߺ���ɫ�ָ�
    Me.lbl����(1).Font.Underline = False
    Me.lbl����(1).ForeColor = lbl����(0).ForeColor
    vsgIllness.FontUnderline = False
    vsgIllness.ForeColor = lbl����(0).ForeColor
End Sub

Private Sub picInfo_Resize()
    On Error Resume Next
    
    lbl˵��.Width = PicInfo.ScaleWidth - lbl˵��.Left * 2
    lbl����(1).Width = PicInfo.ScaleWidth - lbl����(1).Left - lbl˵��.Left
    vsgIllness.Left = lbl����(0).Left
    vsgIllness.Width = PicInfo.ScaleWidth - vsgIllness.Left - lbl˵��.Left
End Sub

Private Sub picList_Resize()
    On Error Resume Next

    rptPath.Left = 0
    rptPath.Top = 0
    rptPath.Width = picList.ScaleWidth
    rptPath.Height = picList.ScaleHeight - vsFile.Height

    vsFile.Left = 0
    vsFile.Top = rptPath.Top + rptPath.Height
    vsFile.Width = picList.ScaleWidth
End Sub

Private Sub rptPath_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim objControl As CommandBarControl
    Dim objRow As ReportRow

    If KeyCode = vbKeyReturn And Shift = 0 Then
        If rptPath.SelectedRows.count > 0 Then
            If Not rptPath.SelectedRows(0).GroupRow Then
                Set objRow = rptPath.SelectedRows(0)
            End If
        End If
        If Not objRow Is Nothing Then
            Set objControl = cbsMain.FindControl(, conMenu_Edit_Modify, True, True)
            If Not objControl Is Nothing Then objControl.Execute
        End If
    End If
End Sub

Private Sub rptPath_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objHitTest As ReportHitTestInfo
    Dim objPopup As CommandBarPopup

    If Button = 2 Then
        Set objHitTest = rptPath.HitTest(X, Y)
        If objHitTest.ht = xtpHitTestReportArea And Not objHitTest.Row Is Nothing Then
            If objHitTest.Row.GroupRow Then
                Set objPopup = cbsMain.FindControl(, conMenu_View_Expend, , True)
            ElseIf objHitTest.Row.Childs.count = 0 Then
                Set objPopup = cbsMain.ActiveMenuBar.Controls(2)
            End If
        End If

        rptPath.SetFocus
        If Not objPopup Is Nothing Then objPopup.CommandBar.ShowPopup
    End If
End Sub

Private Sub rptPath_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim objControl As CommandBarControl

    If Not Row.GroupRow Then
        Set objControl = cbsMain.FindControl(, conMenu_Edit_Modify, True, True)
        If Not objControl Is Nothing Then objControl.Execute
    End If
End Sub

Private Sub rptPath_SelectionChanged()
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, strTmp As String
    Dim arrStr As Variant
    Dim intRowNum As Integer: Dim intColNum As Integer
    Dim i As Long

    On Error GoTo errH

    If rptPath.SelectedRows.count = 0 Then
        Call ClearSubData
    ElseIf rptPath.SelectedRows(0).GroupRow Then
        Call ClearSubData
    Else
        With rptPath.SelectedRows(0)
            lbl˵��.Caption = "˵����" & .Record(COL_˵��).Value

            '��Ӧ������Ϣ
            If .Record(COL_ͨ��).Value = 1 Then
                lbl����(1).Caption = "���ٴ�·�����������������ٴ�����"
            Else
                strSql = "Select B.����,B.���� From ����·������ A,���ű� B Where A.����ID=B.ID And A.·��ID=[1] Order by B.����"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(.Record(COL_ID).Value))
                strTmp = ""
                Do While Not rsTmp.EOF
                    strTmp = strTmp & "," & rsTmp!���� & "-" & rsTmp!����
                    rsTmp.MoveNext
                Loop
                If strTmp <> "" Then
                    lbl����(1).Caption = Mid(strTmp, 2)
                Else
                    lbl����(1).Caption = "<���ٴ�·����δ���������õĿ���>"
                End If
            End If

            '��Ӧ������Ϣ
            vsgIllness.Clear

            strSql = " Select Decode(B.����,NULL,'['||C.����||']'||C.����,'['||B.����||']'||B.����) as ���� ,B.���� " & _
                     " From ����·������ A,��������Ŀ¼ B,�������Ŀ¼ C" & _
                     " Where A.����ID=B.ID(+) And A.���ID=C.ID(+) And A.·��ID=[1] " & _
                     " Order by B.����,C.����"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(.Record(COL_ID).Value))
            strTmp = ""
            mstr�������� = ""
            Do While Not rsTmp.EOF
                strTmp = strTmp & "," & rsTmp!����
                If rsTmp!���� & "" <> "" Then
                    mstr�������� = mstr�������� & "," & rsTmp!����
                End If
                rsTmp.MoveNext
            Loop
            If strTmp <> "" Then
                With vsgIllness
                    arrStr = Split(Mid(strTmp, 2), ",")
                    .Cols = 3: .Rows = ((UBound(arrStr) + 1) + (.Cols - 1)) \ .Cols
                    .Tag = Mid(strTmp, 2)
                    For i = 0 To UBound(arrStr)
                        intRowNum = i \ .Cols
                        intColNum = i Mod .Cols
                        .TextMatrix(intRowNum, intColNum) = arrStr(i)
                    Next i
                    mstr�������� = Mid(mstr��������, 2)
                End With
            Else
                vsgIllness.Rows = 1: vsgIllness.Cols = 1
                vsgIllness.TextMatrix(0, 0) = "<���ٴ�·����δ��������Ӧ�Ĳ���>"
            End If

            '��Ӧ�ļ���Ϣ
            Call LoadPathFile(Val(.Record(COL_ID).Value))

            '·������Ϣ
            Call mfrmContent.zlRefresh(Val(.Record(COL_ID).Value), mstrPrivs, lbl����(1).Caption, vsgIllness.Tag)
        End With

        Call Form_Resize
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function LoadPathFile(ByVal lng·��ID As Long) As Boolean
'���ܣ���ʾ�ٴ�·���ļ����ݺͻ����ٴ�·��
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, i As Long

    On Error GoTo errH

    strSql = "Select �ļ���,������,����ʱ��,��� From ����·���ļ� Where ·��ID=[1] Order by ����ʱ��"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng·��ID)
    With vsFile
        .Rows = .FixedRows '�����
        .Rows = .FixedRows + rsTmp.RecordCount
        .Height = .RowHeight(0) + .RowHeightMin * (.Rows - 1) + Screen.TwipsPerPixelY * 2
        For i = 1 To rsTmp.RecordCount
            .TextMatrix(i, 1) = rsTmp!�ļ���
            Set .Cell(flexcpPicture, i, 0) = zlCommFun.GetFileIcon(rsTmp!�ļ���, True, App.hInstance)
            .Cell(flexcpPictureAlignment, i, 0) = 7

            'ɾ��֮ǰ����ʱ�ļ�
            If gobjFile.FileExists(gobjFile.GetSpecialFolder(TemporaryFolder) & "\" & rsTmp!�ļ���) Then
                gobjFile.DeleteFile gobjFile.GetSpecialFolder(TemporaryFolder) & "\" & rsTmp!�ļ���, True
            End If
            If rsTmp!��� = 1 Then .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &HFF0000
            rsTmp.MoveNext
        Next
    End With

    Call picList_Resize
    LoadPathFile = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ResizeInfoPane()
'���ܣ����ݵ�ǰ��Ϣ���ݣ�������Ϣ�������Ϣ����ߴ��λ��
'˵��������Label��AutoSize�����Զ�������ǩ�߶�
    lbl����(0).Top = lbl˵��.Top + lbl˵��.Height + Screen.TwipsPerPixelY * 6
    lbl����(1).Top = lbl����(0).Top + lbl����(0).Height + Screen.TwipsPerPixelY * 3
    lbl����(0).Top = lbl����(1).Top + lbl����(1).Height + Screen.TwipsPerPixelY * 6

    vsgIllness.Top = lbl����(0).Top + lbl����(0).Height + Screen.TwipsPerPixelY * 3
    '���ݶ�Ӧ����������̬��ʾ��Ӧ������Ϣ�������ʾ5��
    vsgIllness.Height = vsgIllness.RowHeightMin * IIf(vsgIllness.Rows > 5, 5, vsgIllness.Rows)
    vsgIllness.ColWidthMin = vsgIllness.Width / vsgIllness.Cols
    PicInfo.Height = vsgIllness.Top + vsgIllness.Height + Screen.TwipsPerPixelY * 6
End Sub

Private Function RefreshData(Optional ByVal str���� As String, Optional ByVal str���� As String) As Boolean
'���ܣ����ݵ�ǰ���õ�������ȡ�ٴ�·��Ŀ¼����
'���������ڶ�λ
    Dim rsTmp       As ADODB.Recordset
    Dim strSql      As String
    Dim objRecord   As ReportRecord
    Dim objItem     As ReportRecordItem
    Dim objRow As ReportRow, i As Long
    Dim lngPreID As Long, lngPreIdx As Long
    Dim intTypeNum  As Integer
    Dim lngPathColor As Long                'δ���·��Ŀ¼ǰ����ɫֵ

    Screen.MousePointer = 11

    On Error GoTo errH

    'SQL�в��������Ч��,ReportControl��������
    strSql = "Select Distinct a.Id, a.����, a.����, a.����, a.�����Ա�, a.��������, a.˵��, a.ͨ��, a.���°汾, Min(Decode(c.���ʱ��, Null, 0, 1)) As �����" & vbNewLine & _
             "From ����·��Ŀ¼ A, ����·���汾 C" & vbNewLine & _
             "Where a.Id = c.·��id(+)"

    If InStr(mstrPrivs, "ȫԺ·��") = 0 Then
        'û��Ȩ��ʱ��ֻ�ܶ�ֻӦ���ڱ��Ƶ�·�����д���
        strSql = strSql & _
                 " And A.ͨ�� = 2 And Exists" & vbNewLine & _
                 "      (Select 1 From ������Ա C,����·������ D  " & vbNewLine & _
                 "       Where C.��Աid = [1] And D.����id = C.����id And ·��id = A.ID)"
    End If

    strSql = strSql & " Group By a.Id, a.����, a.����, a.����, a.�����Ա�, a.��������, a.˵��, a.ͨ��, a.���°汾 "

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID)

    '��¼����ѡ�еķ���
    If rptPath.SelectedRows.count > 0 Then
        If Not rptPath.SelectedRows(0).GroupRow Then
            lngPreIdx = rptPath.SelectedRows(0).Index    '���ڿ������¶�λ
            lngPreID = rptPath.SelectedRows(0).Record(COL_ID).Value
        End If
    End If

    rptPath.Records.DeleteAll
    Do While Not rsTmp.EOF
        Set objRecord = Me.rptPath.Records.Add()
        Set objItem = objRecord.AddItem(Val(rsTmp!ID))
        Set objItem = objRecord.AddItem("")
        objItem.Icon = img16.ListImages("Path").Index - 1
        Set objItem = objRecord.AddItem("")
        Set objItem = objRecord.AddItem("")
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!����, "<δָ������>")))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!����)))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!����)))
        Set objItem = objRecord.AddItem(CStr(Decode(NVL(rsTmp!�����Ա�, 0), 0, "", 1, "��", 2, "Ů")))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!��������)))
        Set objItem = objRecord.AddItem(CStr("" & rsTmp!˵��))
        Set objItem = objRecord.AddItem(Val(NVL(rsTmp!ͨ��, 1)))
        Set objItem = objRecord.AddItem(Val(NVL(rsTmp!���°汾, 0)))
        Set objItem = objRecord.AddItem(zlCommFun.SpellCode(NVL(rsTmp!����) & "��0"))

        lngPathColor = IIf(Val(rsTmp!�����) = 1, vbBlack, &H80&)
        For i = COL_�к� To COL_ͨ��
            objRecord.Item(i).ForeColor = lngPathColor
        Next

        rsTmp.MoveNext
    Loop

    rptPath.Populate

    '�����ж��ʱ����ʾ�к���
    If rptPath.Rows.count - rptPath.Records.count > 1 Then
        rptPath.Columns(COL_�к�).Visible = True
        rptPath.Columns(COL_�к�).SortAscending = True
    Else
        rptPath.Columns(COL_�к�).Visible = False
    End If

    '�кŸ�ֵ
    For i = 0 To rptPath.Rows.count - 1
        With rptPath.Rows(i)
            If .GroupRow Then intTypeNum = intTypeNum + 1
            If Not .GroupRow Then
                .Record(COL_�к�).Value = i - intTypeNum + 1
            End If
        End With
    Next

    If rptPath.Rows.count = 0 Then
        Call ClearSubData
    Else
        If str���� <> "" And str���� <> "" Then
            For i = 0 To rptPath.Rows.count - 1
                If Not rptPath.Rows(i).GroupRow Then
                    If rptPath.Rows(i).Record(COL_����).Value = str���� _
                       And rptPath.Rows(i).Record(COL_����).Value = str���� Then
                        Set objRow = rptPath.Rows(i): Exit For
                    End If
                End If
            Next
        Else
            If lngPreID <> 0 Then
                '�ȿ��ٶ�λ
                If lngPreIdx <= rptPath.Rows.count - 1 Then
                    If Not rptPath.Rows(lngPreIdx).GroupRow Then
                        If rptPath.Rows(lngPreIdx).Record(COL_ID).Value = lngPreID Then
                            Set objRow = rptPath.Rows(lngPreIdx)
                        End If
                    End If
                End If
                '�ٽ��в���
                If objRow Is Nothing Then
                    For i = 0 To rptPath.Rows.count - 1
                        If Not rptPath.Rows(i).GroupRow Then
                            If rptPath.Rows(i).Record(COL_ID).Value = lngPreID Then
                                Set objRow = rptPath.Rows(i): Exit For
                            End If
                        End If
                    Next
                End If
            End If
            'ȡ��һ���Ƿ�����
            If objRow Is Nothing Then
                For i = 0 To rptPath.Rows.count - 1
                    If Not rptPath.Rows(i).GroupRow Then Set objRow = rptPath.Rows(i): Exit For
                Next
            End If
        End If

        Set rptPath.FocusedRow = objRow    '����ѡ������ʾ�ڿɼ�����,������SelectionChanged�¼�
        Me.stbThis.Panels(2).Text = "���� " & rptPath.Records.count & " ���ٴ�·��"
    End If

    Screen.MousePointer = 0
    RefreshData = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ClearSubData()
'���ԭ�е�������Ϣ
    Dim i As Integer

    lbl˵��.Caption = "˵����"

    lbl����(1).Caption = ""

    vsgIllness.Rows = 0
    vsgIllness.Rows = 5
    vsFile.Rows = vsFile.FixedRows
    vsFile.Height = vsFile.RowHeight(0) + vsFile.RowHeightMin * (vsFile.Rows - 1) + Screen.TwipsPerPixelY * 2

    Me.stbThis.Panels(2).Text = ""

    Call mfrmContent.zlRefresh(0, mstrPrivs, lbl����(1).Caption, vsgIllness.Tag)

    Call Form_Resize
    Call picList_Resize
End Sub

Private Sub txtFind_GotFocus()
    Call zlControl.TxtSelAll(txtFind)
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
'����Enter�����в���
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call FuncFindPath
    End If
End Sub

Private Sub FuncFindPath(Optional ByVal blnNext As Boolean)
'������blnNext=�Ƿ������һ��
    Static blnReStart As Boolean
    Dim blnHave As Boolean, i As Long

    Call zlControl.TxtSelAll(txtFind)
    '��ʼ������
    If rptPath.SelectedRows.count > 0 Then blnHave = True
    If Not blnNext Or blnReStart Or Not blnHave Then
        i = 0    'ReportControl����������0��ʼ
    Else
        i = rptPath.SelectedRows(0).Index + 1
    End If

    '����·��
    For i = i To rptPath.Rows.count - 1
        With rptPath.Rows(i)
            If Not .GroupRow Then
                '��������������жϣ��������ĸ����Ҽ��룬�Ǻ�����������ƣ�������������к�
                If zlCommFun.IsNumOrChar(Trim(txtFind.Text)) Then
                    '��ĸ������
                    If zlCommFun.IsCharAlpha(Trim(txtFind.Text)) Then
                        'ȫ����ĸ����ƴ������
                        If .Record(COL_ƴ������).Value Like "*" & UCase(Trim(txtFind.Text)) & "*" Then
                            Exit For
                        End If
                    Else
                        'ȫ�����ֲ����к�
                        If .Record(COL_�к�).Value Like "*" & Trim(txtFind.Text) & "*" Then
                            Exit For
                        End If
                    End If
                ElseIf zlCommFun.IsCharChinese(Trim(txtFind.Text)) Then
                    '�������� ��������
                    If .Record(COL_����).Value Like "*" & Trim(txtFind.Text) & "*" Then
                        Exit For
                    End If
                End If
            End If
        End With
    Next

    If i <= rptPath.Rows.count - 1 Then
        blnReStart = False
        '����ѡ������ʾ�ڿɼ�����,������SelectionChanged�¼�
        Set rptPath.FocusedRow = rptPath.Rows(i)

        If rptPath.Visible Then rptPath.SetFocus
    Else
        blnReStart = True
        MsgBox IIf(blnNext, "������", "") & "�Ҳ������������������ٴ�·����", vbInformation, gstrSysName
    End If
End Sub

Private Sub txtFind_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'��꾭�����ҿ��ʱ���ֵ���ʾ
    Dim strTip As String

    strTip = "����(Ctrl+F)" & vbCrLf & "������һ��(F3)"
    zlCommFun.ShowTipInfo txtFind.Hwnd, strTip, True
End Sub

Private Sub vsFile_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error Resume Next
    vsFile.ForeColorSel = vsFile.Cell(flexcpForeColor, NewRow, 0)
End Sub

Private Sub vsFile_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If NewCol = 0 Then Cancel = True
End Sub

Private Sub vsFile_DblClick()
'���˫���鿴
    If vsFile.MouseRow >= vsFile.FixedRows Then
        Call vsFile_KeyPress(13)
    End If
End Sub

Private Sub vsFile_KeyDown(KeyCode As Integer, Shift As Integer)
'����Delete��ɾ��
    If KeyCode = vbKeyDelete Then
        If vsFile.Row >= vsFile.FixedRows Then Call FuncPathFileDelete
    End If
End Sub

Private Sub vsFile_KeyPress(KeyAscii As Integer)
'����Enter���鿴
    If KeyAscii = 13 Then
        KeyAscii = 0
        If vsFile.Row >= vsFile.FixedRows Then Call FuncPathFileView
    End If
End Sub

Private Sub vsFile_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'�����Ҽ��˵�
    Dim lngRow As Long
    Dim objPopup As CommandBarPopup

    lngRow = vsFile.MouseRow

    If Button = 2 And lngRow <> -1 Then
        vsFile.SetFocus
        If lngRow <= vsFile.Rows - 1 And lngRow >= vsFile.FixedRows Then
            vsFile.Row = lngRow
        End If

        Set objPopup = cbsMain.FindControl(, conMenu_Edit_Archive, True, True)
        If Not objPopup Is Nothing Then objPopup.CommandBar.ShowPopup
    End If
End Sub

Private Sub FuncExportToXMLBatch()
'���ܣ���������ΪXML�ļ�
    Dim strPath As String, strFile As String
    Dim strFail As String, intCount As Integer
    Dim strMsg As String, i As Long

    If MsgBox("�����ܽ����������ٴ�·������������˰汾��" & _
        vbCrLf & "ÿ�������ļ�����������Ϊ""����-·������.xml""��" & vbCrLf & "����ڵ���Ŀ��λ������ͬ���Ƶ��ļ��򽫱����ǡ�" & _
        vbCrLf & vbCrLf & "Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub

    strPath = zlCommFun.OpenDir(Me.Hwnd, "�����ٴ�·������Ŀ¼", GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "�����ٴ�·��XMLĿ¼"))
    If strPath = "" Then Exit Sub
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "�����ٴ�·��XMLĿ¼", strPath

    Screen.MousePointer = 11
    For i = 0 To rptPath.Records.count - 1
        With rptPath.Records(i)
            Call zlCommFun.ShowFlash(i + 1 & "/" & rptPath.Records.count & "�����ڵ��������ٴ�·��""" & .Item(COL_����).Value & """ ...", Me)
            If .Item(COL_���°汾).Value > 0 Then
                strFile = strPath & "\" & .Item(COL_����).Value & "-" & .Item(COL_����).Value & ".xml"
                If ExportOutPathToXML(.Item(COL_ID).Value, .Item(COL_���°汾).Value, strFile) Then
                    intCount = intCount + 1
                Else
                    strFail = strFail & vbCrLf & strFile
                End If
            End If
        End With
    Next
    Call zlCommFun.StopFlash
    Screen.MousePointer = 0

    strMsg = "������ɣ����ɹ����� " & intCount & " �������ٴ�·���ļ���" & _
        IIf(strFail <> "", "���������ٴ�·������ʧ�ܣ�" & vbCrLf & strFail, "")
    MsgBox strMsg, vbInformation, gstrSysName
End Sub

Private Sub FuncImportFromXMLBatch()
'���ܣ���������XML�ļ�
    Dim arrFile() As String
    Dim strFail As String, strMsg As String
    Dim intCount As Long, i As Long
    Dim intLimit As Integer, strLimit As String, blnLimit As Boolean

    If InStr(mstrPrivs, "������·������") > 0 Then
        intLimit = 0
    ElseIf InStr(mstrPrivs, "30������·��") > 0 Then
        intLimit = 30
    ElseIf InStr(mstrPrivs, "5������·��") > 0 Then
        intLimit = 5
    Else
        MsgBox "������ȷ��Ȩ�����·������������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
        Exit Sub
    End If

    strMsg = "���������������ٴ�·��XML�ļ�ʱ��" & vbCrLf & vbCrLf & _
            "��1.���ϵͳ�в�������ͬ��������Ƶ������ٴ�·�������뽫�����Ӹ������ٴ�·����" & vbCrLf & _
            "��2.���ϵͳ���Ѵ�����ͬ��������Ƶ������ٴ�·�����򽫵��뵽�������ٴ�·���µİ汾�С�" & vbCrLf & _
            "��������������ٴ�·������δ��˵İ汾�����뽫���Ǹð汾�����ݡ�" & vbCrLf & vbCrLf & _
            "Ҫ������"
    If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub

    cdgFile.DialogTitle = "�����ٴ�·��"
    cdgFile.Filter = "XML�ļ�|*.xml"
    cdgFile.Flags = &H200 Or &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
    cdgFile.InitDir = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "�����ٴ�·��XMLĿ¼")
    cdgFile.FileName = ""
    cdgFile.MaxFileSize = 25600 '��ѡʱ�������ļ�������������(Byte)
    cdgFile.CancelError = True
    On Error Resume Next
    cdgFile.ShowOpen
    If Err.Number <> 0 Then
        Err.Clear: Exit Sub
    End If
    On Error GoTo 0

    If InStr(cdgFile.FileName, Chr(0)) > 0 Then
        ReDim arrFile(UBound(Split(cdgFile.FileName, Chr(0))) - 1)
        For i = 0 To UBound(arrFile)
            arrFile(i) = Split(cdgFile.FileName, Chr(0))(0) & "\" & Split(cdgFile.FileName, Chr(0))(i + 1)
        Next
    Else
        ReDim arrFile(0)
        arrFile(0) = cdgFile.FileName
    End If
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "�����ٴ�·��XMLĿ¼", gobjFile.GetParentFolderName(arrFile(0))

    Screen.MousePointer = 11
    For i = 0 To UBound(arrFile)
        Call zlCommFun.ShowFlash(i + 1 & "/" & UBound(arrFile) + 1 & "�����ڵ����ļ�""" & gobjFile.GetFileName(arrFile(i)) & """ ...", Me)
        If ImportOutPathFromXML(arrFile(i), , , intLimit, blnLimit) Then
            intCount = intCount + 1
        Else
            If blnLimit Then
                strLimit = strLimit & vbCrLf & arrFile(i)
            Else
                strFail = strFail & vbCrLf & arrFile(i)
            End If
        End If
    Next
    Call zlCommFun.StopFlash
    Call RefreshData
    Screen.MousePointer = 0

    strMsg = "������ɣ����ɹ����� " & intCount & " �������ٴ�·���ļ���" & _
        IIf(strFail <> "", "���������ٴ�·���ļ�����ʧ�ܣ�" & vbCrLf & strFail, "") & _
        IIf(strLimit <> "", "����·������Ȩ��������δ���룺" & vbCrLf & strLimit, "")
    MsgBox strMsg, vbInformation, gstrSysName
End Sub

Private Sub InitVsgIllness()
'����:��ʼ����Ӧ���ֵ�VSF�ؼ�
    With vsgIllness
        .Cols = 3
        .Rows = 5
        .FixedCols = 0
        .FixedRows = 0
        .AllowSelection = False
        .BackColorBkg = vbWhite
        .RowHeightMin = 300
        .Appearance = flexXPThemes
        .BorderStyle = flexBorderNone
        .ScrollBars = flexScrollBarVertical
        .GridLines = flexGridNone
        .ColWidthMin = .Width / .Cols
    End With
End Sub

Private Sub vsgIllness_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'���֣��������ȥ֮������»��ߣ���ɫ����
    vsgIllness.FontUnderline = True
    vsgIllness.ForeColor = RGB(0, 0, 128)
    vsgIllness.ToolTipText = vsgIllness.Text
End Sub
