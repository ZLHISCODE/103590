VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmUnitInfoEdit 
   BackColor       =   &H80000005&
   Caption         =   "ҽԺ��Ϣά��"
   ClientHeight    =   6450
   ClientLeft      =   6525
   ClientTop       =   3510
   ClientWidth     =   10725
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frmUnitInfoEdit.frx":0000
   ScaleHeight     =   6450
   ScaleWidth      =   10725
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboStation 
      Height          =   300
      Left            =   5640
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   810
      Width           =   1125
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1800
      Left            =   9120
      ScaleHeight     =   1800
      ScaleWidth      =   1800
      TabIndex        =   8
      Top             =   2400
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "ˢ��(&R)"
      Height          =   350
      Left            =   8160
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5880
      Width           =   1100
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   840
      ScaleHeight     =   3855
      ScaleWidth      =   8055
      TabIndex        =   2
      Top             =   1200
      Width           =   8055
      Begin VSFlex8Ctl.VSFlexGrid vsUnitInfo 
         Height          =   3615
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   7800
         _cx             =   13758
         _cy             =   6376
         Appearance      =   1
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
         BackColorSel    =   16761024
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   4
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   2
         RowHeightMin    =   300
         RowHeightMax    =   1500
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmUnitInfoEdit.frx":04F9
         ScrollTrack     =   0   'False
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   1
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   0   'False
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
         Editable        =   2
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
   End
   Begin VB.CommandButton cmdItemsDelete 
      Caption         =   "ɾ����Ŀ(&D)"
      Height          =   350
      Left            =   6600
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton cmdItemsModify 
      Caption         =   "������Ŀ(&M)"
      Height          =   350
      Left            =   5160
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton cmdItemsNew 
      Caption         =   "������Ŀ(&N)"
      Height          =   350
      Left            =   3600
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5880
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog cdgPub 
      Left            =   3000
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblSta 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ժ��"
      Height          =   180
      Left            =   5160
      TabIndex        =   9
      Top             =   870
      Width           =   360
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ṩҽԺ���๫����Ϣ�Ķ�������ݱ༭�Ĺ��ܡ�"
      Height          =   180
      Left            =   960
      TabIndex        =   1
      Top             =   870
      Width           =   3960
   End
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   240
      Picture         =   "frmUnitInfoEdit.frx":05D4
      Top             =   600
      Width           =   480
   End
   Begin VB.Image imgButtonDel 
      Height          =   240
      Left            =   2280
      Picture         =   "frmUnitInfoEdit.frx":0C6A
      Top             =   6000
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgButtonEdit 
      Height          =   240
      Left            =   2520
      Picture         =   "frmUnitInfoEdit.frx":74BC
      Top             =   6000
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ҽԺ��Ϣά��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   195
      TabIndex        =   0
      Top             =   105
      Width           =   1440
   End
   Begin VB.Menu mnuPop 
      Caption         =   "�����˵�"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuPopNew 
         Caption         =   "������Ŀ"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuPopModfy 
         Caption         =   "������Ŀ"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuPopDel 
         Caption         =   "ɾ����Ŀ"
         Shortcut        =   ^D
      End
   End
End
Attribute VB_Name = "frmUnitInfoEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum UnitCol
    Col_���� = 0
    Col_��Ŀ = 1
    Col_�Ƿ�ͼƬ = 2
    Col_���� = 3
    Col_Edit = 4
    Col_Del = 5
    Col_�������� = 6
End Enum
Private mstrStation As String 'վ��
'===========================================================================
'==�����ӿ�
'===========================================================================
Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = False
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'�������ڵ��ã�ʵ�־���Ĵ�ӡ����
'���û�пɴ�ӡ�ģ�������һ���յĽӿ�
End Sub

'===========================================================================
'==�¼�
'===========================================================================
Private Sub cboStation_Click()
    Dim strCurStation As String
    strCurStation = cboStation.ItemData(cboStation.ListIndex)
    If strCurStation = "-1" Then strCurStation = ""
    If strCurStation <> mstrStation And cboStation.Tag <> "" Then
        mstrStation = strCurStation
        Call RefreshData
    End If
End Sub

Private Sub cmdItemsDelete_Click()
    Dim strSQL As String
    Dim strRemarks As String
    
    With vsUnitInfo
        If .TextMatrix(.Row, Col_��������) <> "1" Then
            If MsgBox("ȷ��Ҫɾ��""" & .TextMatrix(.Row, Col_��Ŀ) & """��", vbInformation + vbYesNo, gstrSysName) = vbNo Then
                Exit Sub
            End If
        Else
            If MsgBox("��Ŀ""" & .TextMatrix(.Row, Col_��Ŀ) & """�Ѿ����ܱ�ʹ�ã�ȷ��Ҫɾ����", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
        '��֤��ݲ��������˵��
        If Not CheckAuditStatus("0312", "ɾ����Ŀ", strRemarks) Then Exit Sub
        On Error GoTo ErrH
        strSQL = "Zltools.b_Public.Zlunitinfoitemchange(2,'" & .TextMatrix(.Row, Col_����) & "')"
        Call ExecuteProcedure(strSQL, Me.Caption, gcnOracle)
        '������Ҫ������־
        Call SaveAuditLog(3, "ɾ����Ŀ", .TextMatrix(.Row, Col_��Ŀ), strRemarks)
        .RemoveItem .Row
        Call SetChange
    End With
    Exit Sub
ErrH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub cmdItemsModify_Click()
    Dim intType As Integer, strName As String, strNo As String
    With vsUnitInfo
        strNo = .TextMatrix(.Row, Col_����)
        strName = .TextMatrix(.Row, Col_��Ŀ)
        If frmUnitItemEdit.ShowMe(strNo, strName, intType) Then
            If Val(.TextMatrix(.Row, Col_�Ƿ�ͼƬ)) <> intType Then '�޸������ͣ���������ݱ��Ϊδ�ı�
                .Redraw = flexRDNone
                .TextMatrix(.Row, Col_��������) = "" '��־����������
                .Cell(flexcpData, .Row, Col_����) = "" '���ͼƬ·��
                .TextMatrix(.Row, Col_����) = "" '����ı�����
                Set .Cell(flexcpPicture, .Row, Col_����) = Nothing '���ͼƬ
            End If
            .TextMatrix(.Row, Col_��Ŀ) = strName
            .TextMatrix(.Row, Col_�Ƿ�ͼƬ) = intType
            Call SetChange
        End If
    End With
End Sub

Private Sub cmdItemsNew_Click()
    Dim intType As Integer, strName As String, strNo As String
    With vsUnitInfo
        If frmUnitItemEdit.ShowMe(strNo, strName, intType) Then
            .Rows = .Rows + 1
            .Row = .Rows - 1
            .TextMatrix(.Row, Col_����) = strNo
            .TextMatrix(.Row, Col_��Ŀ) = strName
            .TextMatrix(.Row, Col_�Ƿ�ͼƬ) = intType
            Call SetChange
        End If
    End With
End Sub

Private Sub cmdRefresh_Click()
    Call RefreshData
End Sub

Private Sub Form_Activate()
    picMain.Refresh
    Me.Refresh
End Sub

Private Sub Form_Load()
    Call LoadStation
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    picMain.Height = Me.ScaleHeight - picMain.Top - cmdRefresh.Height - 150
    picMain.Width = Me.ScaleWidth - picMain.Left - 90
    cmdRefresh.Left = Me.ScaleWidth - cmdRefresh.Width - 60
    cmdRefresh.Top = Me.ScaleHeight - cmdRefresh.Height - 90
    Call SetCtrlPosOnLine(False, 0, cmdRefresh, (cmdRefresh.Width + cmdItemsDelete.Width + 60) * -1, cmdItemsDelete, (cmdItemsDelete.Width + cmdItemsModify.Width + 60) * -1, cmdItemsModify, (cmdItemsModify.Width + cmdItemsNew.Width + 60) * -1, cmdItemsNew)
    Call picMain_Resize
End Sub

Private Sub mnuPopDel_Click()
    Call cmdItemsDelete_Click
End Sub

Private Sub mnuPopModfy_Click()
    Call cmdItemsModify_Click
End Sub

Private Sub mnuPopNew_Click()
    Call cmdItemsNew_Click
End Sub

Private Sub picMain_Resize()
    Dim lngWith  As Long, i As Integer
    On Error Resume Next
    With vsUnitInfo
        .Redraw = flexRDNone
        .Height = picMain.ScaleHeight - 15
        .Width = picMain.ScaleWidth - 15
        '��֤�����и����϶��Զ���չ
        For i = 0 To .Cols - 1
            If Not .ColHidden(i) Then
                lngWith = lngWith + .ColWidth(i)
            End If
        Next
        lngWith = .Width - (lngWith - .ColWidth(Col_����))
        .ColWidth(Col_����) = lngWith - 60
        If VScrollVisible(vsUnitInfo) Then
            .ColWidth(Col_����) = .ColWidth(Col_����) - 300
        End If
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub vsUnitInfo_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim i As Integer
    
    cmdItemsModify.Enabled = NewRow > 0
    cmdItemsDelete.Enabled = NewRow > 0
    mnuPopDel.Enabled = NewRow > 0
    mnuPopModfy.Enabled = NewRow > 0
    If NewRow = -1 Or NewCol = -1 Then Exit Sub
    With vsUnitInfo
        .Redraw = flexRDNone
        '���ͼƬ
        For i = .FixedRows To .Rows - 1
            Set .Cell(flexcpPicture, i, Col_Edit) = Nothing
            Set .Cell(flexcpPicture, i, Col_Del) = Nothing
        Next
        .ComboList = ""
        .FocusRect = flexFocusSolid
'        .FocusRect = flexFocusHeavy
        If NewRow >= .FixedRows Then
            Set .CellButtonPicture = Nothing
            '��ʾͼƬ
            If NewCol = Col_Edit Then
                .ComboList = "..."
                .FocusRect = flexFocusNone
                Set .CellButtonPicture = imgButtonEdit.Picture
                If .TextMatrix(NewRow, Col_��������) = "1" Then
                    Set .Cell(flexcpPicture, NewRow, Col_Del) = imgButtonDel.Picture
                End If
            ElseIf NewCol = Col_Del Then
                If .TextMatrix(NewRow, Col_��������) = "1" Then
                    .ComboList = "..."
                    .FocusRect = flexFocusNone
                    Set .CellButtonPicture = imgButtonDel.Picture
                End If
                Set .Cell(flexcpPicture, NewRow, Col_Edit) = imgButtonEdit.Picture
            Else
                If .TextMatrix(NewRow, Col_��������) = "1" Then
                    Set .Cell(flexcpPicture, NewRow, Col_Del) = imgButtonDel.Picture
                End If
                Set .Cell(flexcpPicture, NewRow, Col_Edit) = imgButtonEdit.Picture
            End If
        End If
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub vsUnitInfo_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Col <> Col_��Ŀ And Col <> Col_����
End Sub

Private Sub vsUnitInfo_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim strReturn As String
    Dim vPoint As POINTAPI
    Dim objPic As StdPicture
    Dim mobjEdit As New frmContentEdit
    With vsUnitInfo
        If Col = Col_Del Then
            Call vsUnitInfo_KeyDown(vbKeyDelete, 0)
        ElseIf Col = Col_Edit Then
            If .TextMatrix(.Row, Col_�Ƿ�ͼƬ) = "1" Then
                cdgPub.Filter = "����ͼ���ļ�|*.ico;*.bmp;*.gif;*.jpg|ICON ͼ��(*.ico)|*.ico|λͼͼ��(*.bmp)|*.bmp|GIF ͼ��(*.gif)|*.gif|JPEG ͼ��(*.jpg)|*.jpg|�����ļ�|*.*"
                cdgPub.flags = &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
                cdgPub.InitDir = App.Path
                cdgPub.CancelError = True
                On Error Resume Next
                cdgPub.ShowOpen
                If err.Number <> 0 Then
                    err.Clear: Exit Sub
                End If
                strReturn = cdgPub.FileName
                '����Ƿ�֧��
                On Error Resume Next
                Set objPic = LoadPicture(strReturn)
                If err.Number <> 0 Then
                    MsgBox "��֧�ֵ�ͼ���ʽ��", vbInformation, gstrSysName
                    err.Clear: Exit Sub
                End If
                If strReturn = "" Then Exit Sub
                .Cell(flexcpData, Row, Col_����) = strReturn  '�洢�ļ�·��
                Call SaveData(Row)
                Call RefreshData(Row)
            Else
                strReturn = .TextMatrix(Row, Col_����)
                '��ȡ��ǰλ��
                vPoint = GetCoordPos(.hwnd, .CellLeft - .ColWidth(Col_����), .CellTop + .CellHeight)
                If mobjEdit.ShowMe(frmMDIMain, strReturn, vPoint.x, vPoint.y, , .ColWidth(Col_����)) Then
                    If strReturn <> .TextMatrix(Row, Col_����) Then
                        .TextMatrix(Row, Col_����) = strReturn
                        Call SaveData(Row)
                        Call RefreshData(Row)
                    End If
                End If
                .Col = Col_����
            End If
        End If
    End With
End Sub

Private Sub vsUnitInfo_Click()
    With vsUnitInfo
        If (.MouseCol = Col_Del Or .MouseCol = Col_Edit) And .MouseRow >= .FixedRows Then
            .Select .MouseRow, .MouseCol
            Call vsUnitInfo_CellButtonClick(.MouseRow, .MouseCol)
        End If
    End With
End Sub

Private Sub vsUnitInfo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnHave As Boolean
    If KeyCode = vbKeyDelete Then
        With vsUnitInfo
            If .Row >= .FixedRows Then
                '�ж��Ƿ��������
                If .TextMatrix(.Row, Col_��������) = "1" Then
                    If MsgBox("ȷʵҪ���������Ŀ��Ϣ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        If .TextMatrix(.Row, Col_�Ƿ�ͼƬ) = "1" Then
                            .Cell(flexcpData, .Row, Col_����) = ""  '����ļ�·��
                        Else
                            .TextMatrix(.Row, Col_����) = ""
                        End If
                        Call SaveData(.Row)
                        Call RefreshData(.Row)
                    End If
                End If
            End If
        End With
    End If
End Sub

Private Sub vsUnitInfo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call EnterNextCell
    End If
End Sub

Private Sub vsUnitInfo_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu Me.mnuPop, , picMain.Left + vsUnitInfo.Left + vsUnitInfo.CellLeft, picMain.Top + vsUnitInfo.Top + vsUnitInfo.CellTop + vsUnitInfo.CellHeight
    End If
End Sub

Private Sub vsUnitInfo_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Col = Col_Del And vsUnitInfo.TextMatrix(Row, Col_��������) = "" Or Col = Col_����
End Sub

'===========================================================================
'==˽�з���
'===========================================================================
Private Sub LoadStation()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo ErrH
    cboStation.Clear
    mstrStation = "-999"
    strSQL = "Select ���, ���� From Zlnodelist Order By ���"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    cboStation.AddItem "����"
    cboStation.ItemData(cboStation.NewIndex) = -1
    Do While Not rsTmp.EOF
        cboStation.AddItem rsTmp!���� & ""
        cboStation.ItemData(cboStation.NewIndex) = Val(rsTmp!��� & "")
        rsTmp.MoveNext
    Loop
    cboStation.Tag = "��ʼˢ��"
    cboStation.ListIndex = 0
    Exit Sub
ErrH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Sub

Private Sub RefreshData(Optional ByVal lngCurRow As Long)
'���ܣ����ݼ���
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim i As Integer, lngRow As Long
    Dim strTmp As String, strPreCode As String
    Dim strFile As String, strCode As String
    
    On Error GoTo ErrH
    With vsUnitInfo
        '��ȡ��Ϣ�б��Լ��򵥵���Ϣ����
        If lngCurRow < .FixedRows Then
            strSQL = "Select a.����, a.����, a.�Ƿ�ͼƬ, b.�к�, b.����" & vbNewLine & _
                    "From Zltools.Zlunitinfoitem a, (Select �к�, ����, ��Ŀ From Zltools.Zlreginfo Where Nvl(վ��, '�տ�') = Nvl([2], '�տ�')) b" & vbNewLine & _
                    "Where a.���� = b.��Ŀ(+)" & vbNewLine & _
                    "Order By Lpad(a.����, 3, '0'), b.�к�"
            .Rows = .FixedRows
        Else
            strCode = .TextMatrix(lngCurRow, Col_����)
            .RowHeight(lngCurRow) = .RowHeightMin '������ݣ������Զ��иߣ�һ�λظ�ԭʼ�и�
            strSQL = "Select a.����, a.����, a.�Ƿ�ͼƬ, b.�к�, b.����" & vbNewLine & _
                    "From Zltools.Zlunitinfoitem a, (Select �к�, ����, ��Ŀ From Zltools.Zlreginfo Where Nvl(վ��, '�տ�') = Nvl([2], '�տ�')) b" & vbNewLine & _
                    "Where a.���� = b.��Ŀ(+) and a.����=[1]" & vbNewLine & _
                    "Order By Lpad(a.����, 3, '0'), b.�к�"
        End If
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, strCode, mstrStation)
        '����ˢ�£���Ŀ�����ڣ���ˢ����������
        If lngCurRow >= .FixedRows And rsTmp.RecordCount = 0 Then
            Call RefreshData
            Exit Sub
        End If
        strPreCode = ""
        Do While Not rsTmp.EOF
            If strPreCode <> rsTmp!���� Then
                If strPreCode <> "" And strTmp <> "" Then
                    .TextMatrix(lngRow, Col_����) = strTmp
                    .TextMatrix(lngRow, Col_��������) = "1"  '��ʶ��������
                End If
                If lngCurRow < .FixedRows Then
                    .Rows = .Rows + 1: lngRow = .Rows - 1
                Else
                    lngRow = lngCurRow
                End If
                strTmp = rsTmp!���� & "": strPreCode = rsTmp!����
                .TextMatrix(lngRow, Col_����) = rsTmp!����
                .TextMatrix(lngRow, Col_��Ŀ) = rsTmp!����
                .TextMatrix(lngRow, Col_��������) = ""
                .Cell(flexcpData, .Row, Col_����) = ""
                .TextMatrix(lngRow, Col_�Ƿ�ͼƬ) = Val(rsTmp!�Ƿ�ͼƬ & "")
                '��ͼƬʱ��Ҫ������ȡͼƬ
                If Val(rsTmp!�Ƿ�ͼƬ & "") = 1 Then
                    strFile = gclsBase.ReadLob(gcnOracle, 0, rsTmp!���� & "," & mstrStation)
                    If strFile <> "" Then
                        Set .Cell(flexcpPicture, lngRow, Col_����) = PicDrawPicture(LoadPicture(strFile))
                        .TextMatrix(lngRow, Col_��������) = "1" '��ʶ��������
                    Else
                        Set .Cell(flexcpPicture, lngRow, Col_����) = Nothing
                    End If
                End If
            Else
                '�ı�̫��ʱ����д洢
                strTmp = strTmp & rsTmp!����
            End If
            rsTmp.MoveNext
        Loop
        If strPreCode <> "" And strTmp <> "" Then
            .TextMatrix(lngRow, Col_����) = strTmp
            .TextMatrix(lngRow, Col_��������) = "1"  '��ʶ��������
        End If
    End With
    Call SetChange
    Exit Sub
ErrH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Sub

Private Sub SetChange()
'���ܣ�������Ϣ�ı��״̬
    Dim lngWith  As Long, i As Integer, lngHeight As Long
    
    On Error Resume Next
    With vsUnitInfo
        .Redraw = flexRDNone
'        .Cell(flexcpFontBold, .FixedRows, Col_��Ŀ, .Rows - 1, Col_��Ŀ) = True
'        .Cell(flexcpForeColor, .FixedRows, Col_��Ŀ, .Rows - 1, Col_��Ŀ) = &HD2BDB6
        .Cell(flexcpBackColor, .FixedRows, Col_��Ŀ, .Rows - 1, Col_��Ŀ) = &H8000000F
        '��֤�����и����϶��Զ���չ
        For i = 0 To .Cols - 1
            If Not .ColHidden(i) Then
                lngWith = lngWith + .ColWidth(i)
            End If
        Next
        lngWith = .Width - (lngWith - .ColWidth(Col_����))
        .ColWidth(Col_����) = lngWith
        .AutoSize (Col_����)
        .Redraw = flexRDDirect
        If VScrollVisible(vsUnitInfo) Then
            .Redraw = flexRDNone
            .ColWidth(Col_����) = .ColWidth(Col_����) - 285
            .AutoSize (Col_����)
            .Redraw = flexRDDirect
        End If
        If .Row >= .FixedRows Then
            .TopRow = .Row
            .ShowCell .Row, Col_����
            Call vsUnitInfo_AfterRowColChange(-1, -1, .Row, .Col)
        End If
    End With
End Sub

Private Function SaveData(ByVal lngRow As Long) As Boolean
'���ܣ��������ݱ��档
    Dim i As Integer
    Dim arrSQL() As Variant
    
    On Error GoTo ErrH
    arrSQL = Array()
    With vsUnitInfo
        If lngRow >= .FixedRows Then
            If .TextMatrix(lngRow, Col_�Ƿ�ͼƬ) = "1" Then
                Call gclsBase.GetLobSql(0, .TextMatrix(lngRow, Col_��Ŀ) & "," & mstrStation, .Cell(flexcpData, lngRow, Col_����), arrSQL)
            Else
                Call gclsBase.GetRegInfoSQL(.TextMatrix(lngRow, Col_��Ŀ), .TextMatrix(lngRow, Col_����), mstrStation, arrSQL)
            End If
        End If
    End With
    ShowFlash ("���ڱ������ݣ����Ժ�")
    Call gclsBase.ExecuteProcedureBeach(gcnOracle, arrSQL, Me.Caption)
    ShowFlash ("")
    SaveData = True
    Exit Function
ErrH:
    If 0 = 1 Then
        Resume
    End If
    ShowFlash ("")
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Function

Public Function PicDrawPicture(ByRef objPic As StdPicture) As IPictureDisp
'����ͼƬͬ��������
    picDraw.AutoRedraw = True '������������ȡ������ͼ��
    picDraw.Cls
    picDraw.Width = picDraw.ScaleHeight * (objPic.Width / objPic.Height)
    On Error Resume Next
    picDraw.PaintPicture objPic, 0, 0, picDraw.ScaleWidth, picDraw.ScaleHeight
    Set PicDrawPicture = picDraw.Image
End Function

Public Sub EnterNextCell()
    Dim i As Long, j As Long
    
    With vsUnitInfo
        '����һ��Ԫ��ʼѭ������
        If .Row < .FixedRows Then .Row = .FixedRows
        For i = .Row To .Rows - 1
            For j = IIf(i = .Row, .Col + 1, Col_����) To Col_Del
                If Not .ColHidden(j) Then
                    If j = Col_Del And .TextMatrix(i, Col_��������) = "" Then
                    
                    Else
                        Exit For
                    End If
                End If
            Next
            If j <= Col_Del Then Exit For
        Next
        If i <= .Rows - 1 Then
            .Row = i: .Col = j
            .ShowCell .Row, .Col
        ElseIf i = .Rows And j > Col_Del Then
            Call PressKey(vbKeyTab)
        End If
    End With
End Sub
