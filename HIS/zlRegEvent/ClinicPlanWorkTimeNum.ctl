VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ClinicPlanWorkTimeNum 
   ClientHeight    =   6210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9270
   ScaleHeight     =   6210
   ScaleWidth      =   9270
   Begin VB.PictureBox picFunBack 
      BorderStyle     =   0  'None
      Height          =   350
      Left            =   75
      ScaleHeight     =   345
      ScaleWidth      =   8685
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   8685
      Begin VB.CommandButton cmdFun 
         Caption         =   "���ԤԼ(&R)"
         Height          =   350
         Index           =   4
         Left            =   7410
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   0
         Width           =   1155
      End
      Begin VB.CommandButton cmdFun 
         Caption         =   "ȫ��ԤԼ(&A)"
         Height          =   350
         Index           =   3
         Left            =   6165
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   0
         Width           =   1155
      End
      Begin VB.CommandButton cmdFun 
         Caption         =   "������������(&E)"
         Height          =   350
         Index           =   2
         Left            =   4574
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   1515
      End
      Begin VB.CommandButton cmdFun 
         Caption         =   "���޺ŷֶ�(&N)"
         Height          =   350
         Index           =   1
         Left            =   3135
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmdFun 
         Caption         =   "��Ƶ�ηֶ�(&C)"
         Height          =   350
         Index           =   0
         Left            =   1740
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Width           =   1335
      End
      Begin VB.TextBox txtUpd 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   2
         Text            =   "10"
         Top             =   18
         Width           =   345
      End
      Begin MSComCtl2.UpDown updSkip 
         Height          =   315
         Left            =   1410
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   18
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtUpd"
         BuddyDispid     =   196611
         OrigLeft        =   2580
         OrigTop         =   585
         OrigRight       =   2835
         OrigBottom      =   1200
         Max             =   60
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����Ƶ��(��)"
         Height          =   180
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   85
         Width           =   1080
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsTimeWork 
      Height          =   4830
      Left            =   30
      TabIndex        =   9
      Top             =   540
      Width           =   8880
      _cx             =   15663
      _cy             =   8520
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   16772055
      GridColor       =   12632256
      GridColorFixed  =   12632256
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   2
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"ClinicPlanWorkTimeNum.ctx":0000
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
      Editable        =   2
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
      Begin VB.CommandButton cmdԤԼ 
         Caption         =   "Ԥ"
         Height          =   255
         Left            =   4860
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   210
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.CommandButton cmdɾ�� 
         Caption         =   "ɾ"
         Height          =   255
         Left            =   5520
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   210
         Visible         =   0   'False
         Width           =   345
      End
   End
End
Attribute VB_Name = "ClinicPlanWorkTimeNum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'ȱʡ����ֵ
Private Const m_def_CanReCalic = True

'���Ա���:
Dim m_����Ƶ�� As Integer
Dim m_IsDataChanged As Boolean
Dim m_EditMode As gRegistPlanEditMode
Private m_CanReCalic As Boolean

Private mobj������Ϣ�� As ������Ϣ��
Private mobj�ϰ�ʱ�� As �ϰ�ʱ��
Private mcllFixedSN As Collection
Private mcurDate As Date, mcurNextDate As Date
Private mintPreSelFun As Integer  '�ϴ�ѡ��Ĺ���
Private mblnClickedFunBtn As Boolean '�Ƿ�������ť
'*****************************************************************************************
'VsGrid��Cell���и�˵��
'1.��������ҷ�ʱ��
'  ��0�У�ʱ���,yyyy-mm-dd HH:MM:SS
'  ��>0�У�����У�������
'     a.��һ�У���ţ��洢�Ƿ񿪷�ԤԼ
'     b.�ڶ���:�洢ʱ��Σ��÷ֺŷָ�����ʽ:��ʼʱ��;��ֹʱ�� ,ʱ����yyyy-mm-dd HH:MM:SS��ʾ
'2.����������ҷ�ʱ��
'  ��Mod 2:0-��ʾʱ����У���ʽΪ��ʼʱ��;��ֹʱ�� ,ʱ����yyyy-mm-dd HH:MM:SS��ʾ
'          1-��ʾԤԼ����
'*****************************************************************************************
'�¼�����:
Event DataIsChanged()
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "���û���ӵ�н���Ķ������ͷ���귢����"
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "���û��ƶ����ʱ������"
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "���û���ӵ�н���Ķ����ϰ�����갴ťʱ������"

Event TimeIntervalsChanged(ByVal obj������Ϣ�� As ������Ϣ��, ByVal blnClearUnit As Boolean)
'ȱʡ����ֵ:
Const m_def_����Ƶ�� = 5
Const m_def_IsDataChanged = False
Const m_def_EditMode = 0
Private mblnNotClick As Boolean
Private mblnValiedCanSave As Boolean

Public Function LoadData(ByVal obj������Ϣ�� As ������Ϣ��, ByVal obj�ϰ�ʱ�� As �ϰ�ʱ��, _
    Optional ByVal cllFixedSN As Collection, Optional ByVal blnChanged As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���س��ﰲ��
    '���:obj������Ϣ��-���ﰲ�Ŷ���
    '����:���سɹ�, ����true,���򷵻�False
    '����:���˺�
    '����:2016-01-12 12:46:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    Set mobj������Ϣ�� = obj������Ϣ��
    If mobj������Ϣ�� Is Nothing Then Set mobj������Ϣ�� = New ������Ϣ��
    Set mobj�ϰ�ʱ�� = obj�ϰ�ʱ��
    Set mcllFixedSN = cllFixedSN
    If mcllFixedSN Is Nothing Then Set mcllFixedSN = New Collection
    m_IsDataChanged = blnChanged
    mblnClickedFunBtn = False
    
    mcurDate = Date: mcurNextDate = Date + 1
    Call InitFace
    LoadData = LoadDatatoGrid(obj������Ϣ��)
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub ReCalicWordTime()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���¼��㹤��ʱ���
    '����:���˺�
    '����:2016-01-13 15:54:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    Call cmdFun_Click(mintPreSelFun)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub InitFace()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:���˺�
    '����:2016-01-13 09:52:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    mblnNotClick = True
    If mobj������Ϣ�� Is Nothing Then txtUpd.Text = m_def_����Ƶ��: Exit Sub
    txtUpd.Text = IIf(mobj������Ϣ��.����Ƶ�� = 0, m_def_����Ƶ��, mobj������Ϣ��.����Ƶ��)
    
    picFunBack.Visible = m_CanReCalic And (EditMode = ED_RegistPlan_Edit Or EditMode = ED_RegistPlan_NumLimitModify)
    SetFunVisible mobj������Ϣ��.�Ƿ���ſ���
    Call UserControl_Resize
    mblnNotClick = False
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function LoadDatatoGrid(ByVal obj��ż� As ������Ϣ��) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ݵ�����ؼ�
    '���:obj��ż�-������Ϣ��
    '����:���سɹ�������true,���򷵻�Flase
    '����:���˺�
    '����:2016-01-12 17:49:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln��ſ��� As Boolean
    Dim objColAll As New Collection, obj������Ϣ As ������Ϣ
    
    Err = 0: On Error GoTo Errhand:
    bln��ſ��� = obj��ż�.�Ƿ���ſ���
    If bln��ſ��� And mobj������Ϣ��.�Ƿ��ʱ�� = False Then Exit Function
    
    SetFunVisible bln��ſ���
    'obj��ż��������ʱ���Ⱥ�������򣬲�ȻҪ��
    For Each obj������Ϣ In mobj������Ϣ��
        If mobj������Ϣ��.�Ƿ���ſ��� Then
            objColAll.Add Array(obj������Ϣ.���, obj������Ϣ.��ʼʱ��, obj������Ϣ.��ֹʱ��, _
                IIf(obj������Ϣ.�Ƿ�ԤԼ, 1, 0), IIf(obj������Ϣ.�Ƿ�ͣ��, 1, 0))
        Else
            objColAll.Add Array(obj������Ϣ.���, obj������Ϣ.��ʼʱ��, obj������Ϣ.��ֹʱ��, obj������Ϣ.����)
        End If
    Next
    ShowTimeIntervals bln��ſ���, objColAll
    LoadDatatoGrid = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub SetFunVisible(ByVal blnVisible As Boolean)
    cmdFun(1).Visible = blnVisible
    cmdFun(2).Visible = blnVisible
    cmdFun(3).Visible = blnVisible
    cmdFun(4).Visible = blnVisible
    
    cmdFun(3).Enabled = m_EditMode = ED_RegistPlan_Edit Or m_EditMode = ED_RegistPlan_NumLimitModify
    cmdFun(4).Enabled = m_EditMode = ED_RegistPlan_Edit Or m_EditMode = ED_RegistPlan_NumLimitModify
    If mobj������Ϣ�� Is Nothing Then Exit Sub
    cmdFun(3).Enabled = cmdFun(3).Enabled And mobj������Ϣ��.ԤԼ���� <> 1
    cmdFun(4).Enabled = cmdFun(4).Enabled And mobj������Ϣ��.ԤԼ���� <> 1
End Sub

Private Sub cmdFun_Click(index As Integer)
    Dim strTittle As String
    
    On Error GoTo Errhand
    m_IsDataChanged = True: RaiseEvent DataIsChanged
    If mobj������Ϣ��.�Ƿ��ʱ�� Then
        Select Case index
        Case 0 '��Ƶ�ηֶ�
            If AutoSplitNum(0) = False Then Exit Sub
            mblnClickedFunBtn = True
            mintPreSelFun = index
        Case 1 '���޺����ֶ�
            If AutoSplitNum(1) = False Then Exit Sub
            mblnClickedFunBtn = True
            mintPreSelFun = index
        Case 2 '��������
            If AutoSplitNum(2) = False Then Exit Sub
            mblnClickedFunBtn = True
            mintPreSelFun = index
        Case 3   'ȫ��ԤԼ
            Call SetԤԼ��־(False)
        Case 4   'ȡ��ԤԼ
            Call SetԤԼ��־(True)
        Case Else
        End Select
    End If
    RaiseEvent TimeIntervalsChanged(Get������Ϣ��, False)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function AutoSplitNum(ByVal bytType As Byte) As Boolean
    'ʱ��ηֶ�
    '��Σ�
    '   bytType 0-��Ƶ�ηֶ�,1-���޺����ֶ�,2-��������
    Dim objFrmClinicWorkTimeOther As New frmClinicWorkTimeOther
    Dim colNumAll As Collection, i As Integer, k As Integer
    Dim objCol As Collection, varTimes As Variant, varTemp As Variant
    Dim varTime As Variant, intInterval As Integer
    Dim dtStartDate As Date, dtEndDate As Date
    Dim bln��ʱ�� As Boolean, bln��ſ��� As Boolean, lng�޺��� As Long, lng��Լ�� As Long
    Dim str��ʼʱ�� As String, str��ֹʱ�� As String, str��Ϣʱ�� As String
    Dim lngԤ��ʱ�� As Long, dtStart As Date, dtEnd As Date
    Dim dtCurStart As Date, dtCurEnd As Date, dtCur As Date
    Dim lngCount As Long, lngOverplus As Long, lngCurSN As Long
    Dim colTemp As Collection
    
    Err = 0: On Error GoTo Errhand:
    bln��ʱ�� = True: bln��ſ��� = True
    If Not mobj������Ϣ�� Is Nothing Then
        bln��ʱ�� = mobj������Ϣ��.�Ƿ��ʱ��
        bln��ſ��� = mobj������Ϣ��.�Ƿ���ſ���
        If Not mobj�ϰ�ʱ�� Is Nothing Then
            str��ʼʱ�� = mobj�ϰ�ʱ��.��ʼʱ��
            str��ֹʱ�� = mobj�ϰ�ʱ��.����ʱ��
            str��Ϣʱ�� = mobj�ϰ�ʱ��.��Ϣʱ��
            lngԤ��ʱ�� = mobj�ϰ�ʱ��.����Ԥ��ʱ��
        End If
        lng�޺��� = mobj������Ϣ��.�޺���
        lng��Լ�� = IIf(mobj������Ϣ��.ԤԼ���� = 1, 0, _
            IIf(mobj������Ϣ��.��Լ�� = 0 And mobj������Ϣ��.�޺��� <> 0, mobj������Ϣ��.�޺���, mobj������Ϣ��.��Լ��))
    End If
    
    If bln��ſ��� And lng�޺��� <= 0 Then
        MsgBox "�޺���δ���ã����������޺�����", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If str��ʼʱ�� = "" Then
        dtStartDate = Format(mcurDate, "yyyy-mm-dd") & " 01:00:00"
        dtEndDate = Format(mcurDate, "yyyy-mm-dd") & " 23:59:59"
    Else
        dtStartDate = Format(mcurDate, "yyyy-mm-dd") & " " & Format(str��ʼʱ��, "HH:MM")
        dtEndDate = GetWorkTrueDate(dtStartDate, str��ֹʱ��)
    End If
    
    '��ȥԤ��ʱ��
    Call ��ȥԤ��ʱ��(dtStartDate, dtEndDate, lngԤ��ʱ��, str��Ϣʱ��)
    
    If Val(txtUpd.Text) = 0 And bytType <> 2 Then
        MsgBox "Ƶ��δ���ã����ܷ�ʱ�Σ�", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    Set colNumAll = New Collection
    Select Case bytType
    Case 0 '��Ƶ�ηֶ�
        Set colNumAll = CalculatTimeInterval(0, bln��ſ���, Val(txtUpd.Text), lng�޺���, dtStartDate, dtEndDate, str��Ϣʱ��)
    Case 1 '���޺����ֶ�
        Set colNumAll = CalculatTimeInterval(1, bln��ſ���, Val(txtUpd.Text), lng�޺���, dtStartDate, dtEndDate, str��Ϣʱ��)
    Case 2 '��������
        If bln��ſ��� = False Then AutoSplitNum = True: Exit Function
        If objFrmClinicWorkTimeOther.ShowMe(Me, Val(txtUpd.Text), dtStartDate, dtEndDate, str��Ϣʱ��, varTimes) = False Then Exit Function
        If varTimes(0) = "ʱ����" Then
            Set colNumAll = CalculatTimeInterval(0, bln��ſ���, Val(varTimes(1)), lng�޺���, dtStartDate, dtEndDate)
        ElseIf varTimes(0) = "�ֶμ��" Then
            varTemp = Split(varTimes(1), ";")
            For i = 0 To UBound(varTemp)
                varTime = Split(varTemp(i), ",")(0): intInterval = Val(Split(varTemp(i), ",")(1))
                Set objCol = CalculatTimeInterval(0, bln��ſ���, intInterval, lng�޺���, Split(varTime, "��")(0), Split(varTime, "��")(1), "", colNumAll.Count + 1)
                Set colNumAll = AddRange(colNumAll, objCol)
            Next
        End If
    End Select
    
    '���µ���ԤԼ����
    If bln��ʱ�� And bln��ſ��� = False Then
        Set colTemp = New Collection
        intInterval = lng��Լ�� \ colNumAll.Count 'ÿ��ʱ�ε�ƽ����Լ��
        lngOverplus = lng��Լ�� - intInterval * colNumAll.Count 'ʣ��δ���������Լ���������䵽ǰ��������
        For i = 1 To colNumAll.Count
            'Array(���,��ʼʱ��,��ֹʱ��,ԤԼ����)
            If intInterval = 0 Then
                'ƽ����Լ��������ʱ����ǰ���䣬�������ķ��ں���
                colTemp.Add Array(colNumAll(i)(0), colNumAll(i)(1), colNumAll(i)(2), IIf(i <= lngOverplus, 1, 0)), "K_" & i
            Else
                colTemp.Add Array(colNumAll(i)(0), colNumAll(i)(1), colNumAll(i)(2), _
                    intInterval + IIf(i > colNumAll.Count - lngOverplus, 1, 0)), "K_" & i
            End If
        Next
        Set colNumAll = colTemp
    ElseIf lng��Լ�� > 0 Then
        Set colTemp = New Collection
        For i = 1 To colNumAll.Count
            'Array(���,��ʼʱ��,��ֹʱ��,ԤԼ����)
            colTemp.Add Array(colNumAll(i)(0), colNumAll(i)(1), colNumAll(i)(2), 1), "K_" & i
        Next
        Set colNumAll = colTemp
    End If
    
    ShowTimeIntervals bln��ſ���, colNumAll
    AutoSplitNum = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub ShowTimeIntervals(ByVal bln��ſ��� As Boolean, ByVal objCol As Collection)
    '��ʾ����
    '��Σ�
    '   bln��ſ��ƣ�True-��ſ��ƣ�False-����ſ���
    '   objCol:Array(���,��ʼʱ��,��ֹʱ��,�Ƿ�����ԤԼ/��������,�Ƿ�ͣ��)
    Dim varItem As Variant, varTemp As Variant
    Dim i As Integer, j As Integer, blnFind As Boolean
    Dim lngRow As Long, lngCol As Long, strCurTime As String
    
    Err = 0: On Error GoTo Errhand:
    With vsTimeWork
        .Clear
        .Rows = 0: .Cols = 0
        If objCol Is Nothing Then Exit Sub
        If objCol.Count = 0 Then Exit Sub
        .Redraw = flexRDNone
        If bln��ſ��� Then
            .Rows = 2: .Cols = 2
            .FixedRows = 0: .FixedCols = 1
            .MergeCellsFixed = flexMergeRestrictColumns
            .HighLight = flexHighlightAlways
            .AllowSelection = True
            .MergeCol(0) = True
            lngRow = -2: lngCol = 1: strCurTime = ""
            For Each varItem In objCol
                If strCurTime <> Format(varItem(1), "hh:00") Then
                    strCurTime = Format(varItem(1), "hh:00")
                    lngRow = lngRow + 2: lngCol = 1
                    If lngRow > .Rows - 2 Then .Rows = .Rows + 2
                    .TextMatrix(lngRow, 0) = Format(varItem(1), "hh:00")
                    .TextMatrix(lngRow + 1, 0) = Format(varItem(1), "hh:00")
                End If
                If lngCol > .Cols - 1 Then .Cols = .Cols + 1
                .TextMatrix(lngRow, lngCol) = varItem(0)
                .TextMatrix(lngRow + 1, lngCol) = Format(varItem(1), "hh:mm") & "-" & Format(varItem(2), "hh:mm")
                .Cell(flexcpData, lngRow + 1, lngCol) = Format(varItem(1), "yyyy-mm-dd hh:mm:ss") & "��" & Format(varItem(2), "yyyy-mm-dd hh:mm:ss") '�洢ʱ�䷶Χ
                If Val(varItem(3)) = 1 Then '�Ƿ�ԤԼ
                    .Cell(flexcpData, lngRow, lngCol) = 1
                    .Cell(flexcpForeColor, lngRow, lngCol, lngRow + 1, lngCol) = vbBlue
                    .Cell(flexcpFontBold, lngRow, lngCol, lngRow + 1, lngCol) = True
                End If
                If UBound(varItem) = 4 Then
                    If Val(varItem(4)) = 1 Then '�Ƿ�ͣ��
                        .Cell(flexcpBackColor, lngRow, lngCol, lngRow + 1, lngCol) = vbRed
                    End If
                End If
                lngCol = lngCol + 1
            Next
            .Cell(flexcpAlignment, 0, 1, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
            .Cell(flexcpFontSize, 0, 0, .Rows - 1, 0) = 12
            .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
            .Cell(flexcpAlignment, 0, 0, .Rows - 1, 0) = flexAlignCenterTop
            .ColWidth(-1) = 1300: .ColWidth(0) = 800
        Else
            .Clear
            .Rows = 1: .Cols = 8
            .FixedRows = 1: .FixedCols = 0
            .MergeCellsFixed = flexMergeNever
            .HighLight = flexHighlightNever
            .AllowSelection = False
            
            .Editable = IIf(m_EditMode = ED_RegistPlan_Edit Or m_EditMode = ED_RegistPlan_NumLimitModify, flexEDKbdMouse, flexEDNone)
            If Not mobj������Ϣ�� Is Nothing Then
                .Editable = IIf(.Editable = flexEDKbdMouse And mobj������Ϣ��.ԤԼ���� <> 1, flexEDKbdMouse, flexEDNone)
            End If
            For i = 0 To .Cols - 1 Step 2
                .Cell(flexcpText, 0, i, 0, i + 1) = "ʱ���" & vbTab & "ԤԼ����"
            Next
            lngCol = 0: lngRow = 1
            For Each varItem In objCol
                If lngCol > .Cols - 1 Then lngRow = lngRow + 1: lngCol = 0
                If lngRow > .Rows - 1 Then .Rows = .Rows + 1
                .TextMatrix(lngRow, lngCol) = Format(varItem(1), "hh:mm") & "-" & Format(varItem(2), "hh:mm")
                .Cell(flexcpData, lngRow, lngCol) = varItem(0)
                .Cell(flexcpData, lngRow, lngCol + 1) = Format(varItem(1), "yyyy-mm-dd hh:mm:ss") & "��" & Format(varItem(2), "yyyy-mm-dd hh:mm:ss") '�洢ʱ�䷶Χ
                .TextMatrix(lngRow, lngCol + 1) = Val(varItem(3))
                lngCol = lngCol + 2
            Next
            .ColWidth(-1) = 1200
            .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        End If
        .Redraw = flexRDBuffered
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetԤԼ��־(ByVal blnClear As Boolean, Optional lngRow As Long = -1, Optional lngCol As Long = -1, _
    Optional ByVal blnIgnoreErr As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ԤԼ
    '���:blnClear-���ԤԼ
    '     lngRow=-1��lngCol=-1 ������н�������
    '����:���˺�
    '����:2016-01-13 14:50:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng��Լ As Long, lngCount As Long
    Dim lngSum As Long, lng��ԤԼ As Long
    
    Err = 0: On Error GoTo Errhand:
    If Not mobj������Ϣ�� Is Nothing Then
        lng��Լ = mobj������Ϣ��.��Լ��
    End If
    If blnClear = False And lng��Լ = 0 Then Exit Sub
    
    With vsTimeWork
        If lngRow < 0 Or lngCol <= 0 Then
            For lngRow = 0 To .Rows - 1 Step 2
                If blnClear = False Then
                    For lngCol = 1 To .Cols - 1
                        If .TextMatrix(lngRow, lngCol) <> "" And Val(.Cell(flexcpData, lngRow, lngCol)) = 0 Then
                            .Cell(flexcpData, lngRow, 1, lngRow, lngCol) = 1
                            .Cell(flexcpForeColor, lngRow, 1, lngRow + 1, lngCol) = IIf(blnClear, &H80000008, vbBlue)
                            .Cell(flexcpFontBold, lngRow, 1, lngRow + 1, lngCol) = IIf(blnClear, False, True)
                        End If
                    Next
                Else
                    For lngCol = 1 To .Cols - 1
                        lng��ԤԼ = 0
                        Call ValiedCanModify(Val(.TextMatrix(lngRow, lngCol)), 0, False, lng��ԤԼ)
                        blnClear = lng��ԤԼ = 0
                        .Cell(flexcpData, lngRow, lngCol, lngRow, lngCol) = IIf(blnClear, 0, 1)
                        .Cell(flexcpForeColor, lngRow, lngCol, lngRow + 1, lngCol) = IIf(blnClear, &H80000008, vbBlue)
                        .Cell(flexcpFontBold, lngRow, lngCol, lngRow + 1, lngCol) = IIf(blnClear, False, True)
                    Next
                End If
            Next
        Else
'            lngSum = GetԤԼ����
            If lngRow Mod 2 = 1 Then lngRow = lngRow - 1
'            If lngSum + 1 > lng��Լ And blnClear = False And Val(.Cell(flexcpData, lngRow, lngCol)) <> 1 Then
'                If blnIgnoreErr = False Then MsgBox "������Լ��" & lng��Լ & "�����������ã�", vbInformation + vbOKOnly, gstrSysName
'                Exit Sub
'            End If
            .Cell(flexcpData, lngRow, lngCol) = IIf(blnClear, 0, 1)
            .Cell(flexcpForeColor, lngRow, lngCol, lngRow + 1, lngCol) = IIf(blnClear, &H80000008, vbBlue)
            .Cell(flexcpFontBold, lngRow, lngCol, lngRow + 1, lngCol) = IIf(blnClear, False, True)
        End If
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function GetԤԼ����() As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡԤԼ����
    '����:����ԤԼ����
    '����:���˺�
    '����:2016-01-13 15:04:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, lngCol As Long
    Dim lngSum As Long
    With vsTimeWork
        lngSum = 0
        For lngRow = 0 To .Rows - 1 Step 2
            For lngCol = 1 To .Cols - 1
                If Val(.Cell(flexcpData, lngRow, lngCol)) = 1 Then lngSum = lngSum + 1
            Next
        Next
    End With
    GetԤԼ���� = lngSum
End Function

Private Sub cmdFun_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdɾ��_Click()
    Dim lngRow As Long
    
    On Error GoTo Errhand
    m_IsDataChanged = True: RaiseEvent DataIsChanged
    With vsTimeWork
        lngRow = .Row
        If lngRow Mod 2 = 1 Then lngRow = lngRow - 1
        If .TextMatrix(lngRow, .Col) = "" Then cmdɾ��.Visible = False: Exit Sub
        If ValiedCanModify(Val(.TextMatrix(lngRow, .Col)), 0, True) = False Then
            MsgBox "��ǰʱ�λ�ǰʱ��֮���ʱ���ѱ�ʹ�ã�����ɾ����", vbInformation, gstrSysName
            Exit Sub
        End If
        Call DeleteTime(.Row, .Col)
    End With
    RaiseEvent TimeIntervalsChanged(Get������Ϣ��, False)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdɾ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdԤԼ_Click()
    Dim blnClear As Boolean
    Dim lngRow As Long
    Dim i As Long, j As Long
    Dim intStartRow As Integer, intEndRow As Integer
    Dim intStartCol As Integer, intEndCol As Integer
    
    On Error GoTo Errhand
    m_IsDataChanged = True: RaiseEvent DataIsChanged
    With vsTimeWork
        If .Row < 0 Or .Col < 0 Then Exit Sub
        If .Row = .RowSel And .Col = .ColSel Then
            lngRow = .Row
            If lngRow Mod 2 = 1 Then lngRow = lngRow - 1
            If .TextMatrix(lngRow, .Col) = "" Then cmdԤԼ.Visible = False: Exit Sub
            If ValiedCanModify(Val(.TextMatrix(lngRow, .Col)), 0) = False Then
                MsgBox "��ǰʱ���ѱ�ʹ�ã����ܵ�����", vbInformation, gstrSysName
                Exit Sub
            End If
            blnClear = Val(.Cell(flexcpData, lngRow, .Col)) = 1
            Call SetԤԼ��־(blnClear, lngRow, .Col)
        Else
            '82227����������
            intStartRow = IIf(.Row > .RowSel, .RowSel, .Row)
            intEndRow = IIf(.Row > .RowSel, .Row, .RowSel)
            intStartCol = IIf(.Col > .ColSel, .ColSel, .Col)
            intEndCol = IIf(.Col > .ColSel, .Col, .ColSel)
            For i = intStartRow To intEndRow Step 2
                For j = intStartCol To intEndCol
                    If .TextMatrix(i, j) <> "" And ValiedCanModify(Val(.TextMatrix(i - (i Mod 2), j)), 0) Then
                        blnClear = Val(.Cell(flexcpData, i - (i Mod 2), j)) = 1
                        Call SetԤԼ��־(blnClear, i, j, True)
                    End If
                Next
            Next
            .Select intStartRow, intStartCol
        End If
    End With
    RaiseEvent TimeIntervalsChanged(Get������Ϣ��, False)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdԤԼ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtUpd_Change()
    If mblnNotClick Then Exit Sub
    m_IsDataChanged = True: RaiseEvent DataIsChanged
End Sub

Private Sub txtUpd_GotFocus()
    zlControl.TxtSelAll txtUpd
End Sub

Private Sub txtUpd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub txtUpd_LostFocus()
    If mobj������Ϣ�� Is Nothing Then Exit Sub
    mobj������Ϣ��.����Ƶ�� = Val(txtUpd.Text)
End Sub

Private Sub txtUpd_Validate(Cancel As Boolean)
    If Val(txtUpd.Text) > 60 Or Val(txtUpd.Text) < 1 Then
        MsgBox "����Ƶ�β��ܴ���60���ӻ�С��1���ӣ�", vbInformation, gstrSysName
        zlControl.TxtSelAll txtUpd
        Cancel = True: Exit Sub
    End If
End Sub

Private Sub updSkip_Change()
    If mblnNotClick Then Exit Sub
    m_IsDataChanged = True: RaiseEvent DataIsChanged
    mobj������Ϣ��.����Ƶ�� = Val(txtUpd.Text)
End Sub

Private Sub UserControl_Initialize()
    mcurDate = Date: mcurNextDate = Date + 1
End Sub

Private Sub UserControl_LostFocus()
    cmdɾ��.Visible = False
End Sub

Private Sub UserControl_Resize()
    Err = 0: On Error Resume Next
    With UserControl
        vsTimeWork.Left = .ScaleLeft
        vsTimeWork.Top = IIf(m_CanReCalic And (EditMode = ED_RegistPlan_Edit Or EditMode = ED_RegistPlan_NumLimitModify), picFunBack.Top + picFunBack.Height, 0) + 30
        vsTimeWork.Height = .ScaleHeight - vsTimeWork.Top
        vsTimeWork.Width = .ScaleWidth
    End With
End Sub

Private Function CheckAutoSplitDateIsValied(ByVal dtCurdate As Date) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Զ����������Ƿ�Ϸ�
    '���:dtCurDate-��ǰ����
    '����:�Ϸ�����true,���򷵻�False
    '����:���˺�
    '����:2016-01-13 11:00:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, varTemp As Variant
    Dim i As Long, str��ʼʱ�� As String, str����ʱ�� As String
    Dim dtStartDate As Date, dtEndDate As Date
    
    Err = 0: On Error GoTo Errhand:
    CheckAutoSplitDateIsValied = True
    If mobj������Ϣ�� Is Nothing Then Exit Function
    If mobj�ϰ�ʱ��.��Ϣʱ�� = "" Then Exit Function
    
    varData = Split(mobj�ϰ�ʱ��.��Ϣʱ��, ";")
    For i = 0 To UBound(varData)
        If varData(i) <> "" Then
            varTemp = Split(varData(i), "-")
            If UBound(varTemp) <> 0 Then
                str��ʼʱ�� = varTemp(0)
                str����ʱ�� = varTemp(1)
                If CDate(str��ʼʱ��) > CDate(str����ʱ��) Then
                    dtStartDate = CDate(Format(mcurDate, "yyyy-mm-dd") & " " & str��ʼʱ�� & ":00")
                    dtEndDate = CDate(Format(mcurNextDate, "yyyy-mm-dd") & " " & str����ʱ�� & ":59")
                Else
                    dtStartDate = CDate(Format(mcurDate, "yyyy-mm-dd") & " " & str��ʼʱ�� & ":00")
                    dtEndDate = CDate(Format(mcurDate, "yyyy-mm-dd") & " " & str����ʱ�� & ":59")
                End If
                If dtCurdate >= dtStartDate And dtCurdate <= dtEndDate Then
                    CheckAutoSplitDateIsValied = False: Exit Function
                End If
            End If
        End If
    Next
    CheckAutoSplitDateIsValied = True
    Exit Function
Errhand:
    CheckAutoSplitDateIsValied = True
End Function

Private Function Get������Ϣ��() As ������Ϣ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ��
    '����:������Ϣ��
    '����:���˺�
    '����:2016-01-13 12:34:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, lngCol As Integer
    Dim objNums As New ������Ϣ��, objNum As ������Ϣ, blnԤԼ As Boolean
    Dim lngSum As Long, lngNO As Long, varTemp As Variant
    Dim i As Long
    Dim blnFind As Boolean '����ʱ����ſ���ʱ���Ƿ������ÿ�ԤԼʱ�Σ����һ����û���õĻ���Ĭ������ʱ�ζ���ԤԼ
    
    Err = 0: On Error GoTo Errhand:
    
    '����δ�ı䣬ֱ�ӷ���ԭ���ϵĸ���
    If m_IsDataChanged = False Then
        Set Get������Ϣ�� = mobj������Ϣ��.Clone
        Exit Function
    End If
    
    '�����Ѹı䣬���¹��켯�϶���
    Set objNums = mobj������Ϣ��.Clone
    objNums.RemoveAll
    objNums.�Ƿ��޸� = True
    If objNums.�Ƿ��ʱ�� And Not objNums.�Ƿ���ſ��� And vsTimeWork.FixedCols = 0 Then
        For lngRow = 1 To vsTimeWork.Rows - 1
            For lngCol = 0 To vsTimeWork.Cols - 1 Step 2
               If vsTimeWork.TextMatrix(lngRow, lngCol) <> "" Then
                    lngNO = Val(vsTimeWork.Cell(flexcpData, lngRow, lngCol))
                    varTemp = Split(vsTimeWork.Cell(flexcpData, lngRow, lngCol + 1), "��")
                    lngSum = Val(vsTimeWork.TextMatrix(lngRow, lngCol + 1))
                    Set objNum = New ������Ϣ
                    With objNum
                        .��� = lngNO
                        .��ʼʱ�� = varTemp(0)
                        .��ֹʱ�� = varTemp(1)
                        .���� = lngSum
                        .�Ƿ�ԤԼ = True
                    End With
                    objNums.AddItem objNum
               End If
            Next
        Next
    ElseIf objNums.�Ƿ��ʱ�� And objNums.�Ƿ���ſ��� And vsTimeWork.FixedCols = 1 Then
        For lngRow = 0 To vsTimeWork.Rows - 1 Step 2
            For lngCol = 1 To vsTimeWork.Cols - 1
                If vsTimeWork.TextMatrix(lngRow, lngCol) <> "" _
                    And vsTimeWork.Cell(flexcpFontStrikethru, lngRow, lngCol) = False Then '��ɾ���ߵı�ʾ����Ҫɾ����
                    
                    lngNO = Val(vsTimeWork.TextMatrix(lngRow, lngCol))
                    varTemp = Split(vsTimeWork.Cell(flexcpData, lngRow + 1, lngCol), "��")
                    blnԤԼ = Val(vsTimeWork.Cell(flexcpData, lngRow, lngCol)) = 1
                    If blnԤԼ Then blnFind = True
                    Set objNum = New ������Ϣ
                    With objNum
                        .��� = lngNO
                        .��ʼʱ�� = varTemp(0)
                        .��ֹʱ�� = varTemp(1)
                        .���� = 1
                        .�Ƿ�ԤԼ = blnԤԼ
                    End With
                    objNums.AddItem objNum
                End If
            Next
        Next
'        If blnFind = False And mobj������Ϣ��.ԤԼ���� <> 1 Then
'            'ȫ������ԤԼ
'            For i = 1 To objNums.Count
'                objNums(i).�Ƿ�ԤԼ = True
'            Next
'        End If
    ElseIf objNums.�Ƿ��ʱ�� = False And objNums.�Ƿ���ſ��� Then '������Ų���ʱ�ε��Զ��������
        For i = 1 To objNums.�޺���
            Set objNum = New ������Ϣ
            With objNum
                .��� = i
                .���� = 1
                .�Ƿ�ԤԼ = True '������ԤԼ
                'ʱ�䷶Χ��дΪʱ��εĿ�ʼʱ�����ֹʱ��
                If Not mobj�ϰ�ʱ�� Is Nothing Then
                    .��ʼʱ�� = mobj�ϰ�ʱ��.��ʼʱ��
                    .��ֹʱ�� = mobj�ϰ�ʱ��.����ʱ��
                End If
            End With
            objNums.AddItem objNum
        Next
    End If
    Set Get������Ϣ�� = objNums
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=14,0,0,0
Public Property Get Get����() As ������Ϣ��
   Set Get���� = Get������Ϣ��
End Property

Public Property Get �޺���() As Long
    If mobj������Ϣ�� Is Nothing Then Set mobj������Ϣ�� = New ������Ϣ��
    �޺��� = mobj������Ϣ��.�޺���
End Property

Public Property Let �޺���(ByVal vNewValue As Long)
    Dim lngOld As Long
    
    On Error GoTo Errhand
    If mobj������Ϣ�� Is Nothing Then Set mobj������Ϣ�� = New ������Ϣ��
    lngOld = mobj������Ϣ��.�޺���
    mobj������Ϣ��.�޺��� = vNewValue
    
    If mobj������Ϣ��.�޺��� = lngOld Then Exit Property
    
    m_IsDataChanged = True: RaiseEvent DataIsChanged
    If mobj������Ϣ��.�޺��� = 0 Then
        ShowTimeIntervals True, Nothing
        RaiseEvent TimeIntervalsChanged(Get������Ϣ��, False)
        Exit Property
    End If
    
    '���¼���ʱ��
    Call cmdFun_Click(mintPreSelFun)
    Exit Property
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Property

Public Property Get ��Լ��() As Long
   If mobj������Ϣ�� Is Nothing Then Set mobj������Ϣ�� = New ������Ϣ��
   ��Լ�� = mobj������Ϣ��.��Լ��
End Property

Public Property Let ��Լ��(ByVal vNewValue As Long)
    Dim lngOld As Long, lng��ԤԼ As Long
    Dim lngRow As Long, lngCol As Long, lngSum As Long
    
    If mobj������Ϣ�� Is Nothing Then Set mobj������Ϣ�� = New ������Ϣ��
    lngOld = mobj������Ϣ��.��Լ��
    mobj������Ϣ��.��Լ�� = vNewValue
    
    If mobj������Ϣ��.��Լ�� = lngOld Then Exit Property
    
    On Error GoTo Errhand
    m_IsDataChanged = True: RaiseEvent DataIsChanged
    lngSum = mobj������Ϣ��.��Լ��
    If mobj������Ϣ��.�Ƿ��ʱ�� Then
        If mobj������Ϣ��.�Ƿ���ſ��� Then '��ʱ�Σ���ſ���
            If mobj������Ϣ��.��Լ�� = 0 Then
                'ȫ��ȡ��ԤԼ
                Call SetԤԼ��־(True)
            End If
        ElseIf mobj������Ϣ��.ԤԼ���� = 1 Then '��ʱ�Σ�����ſ���
            For lngRow = 1 To vsTimeWork.Rows - 1
                For lngCol = 0 To vsTimeWork.Cols - 1 Step 2
                    If vsTimeWork.TextMatrix(lngRow, lngCol) <> "" Then
                        lng��ԤԼ = 0
                        Call ValiedCanModify(Val(vsTimeWork.Cell(flexcpData, lngRow, lngCol)), 0, False, lng��ԤԼ)
                        vsTimeWork.TextMatrix(lngRow, lngCol + 1) = lng��ԤԼ
                    End If
                Next
            Next
        End If
    End If
    RaiseEvent TimeIntervalsChanged(Get������Ϣ��, False)
    Exit Property
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Property

Public Property Get ������ſ���() As Boolean
    If mobj������Ϣ�� Is Nothing Then Set mobj������Ϣ�� = New ������Ϣ��
    ������ſ��� = mobj������Ϣ��.�Ƿ���ſ���
End Property

Public Property Let ������ſ���(ByVal vNewValue As Boolean)
    If mobj������Ϣ�� Is Nothing Then Set mobj������Ϣ�� = New ������Ϣ��
    mobj������Ϣ��.�Ƿ���ſ��� = vNewValue
    
    m_IsDataChanged = True: RaiseEvent DataIsChanged
    
    SetFunVisible vNewValue
    ShowTimeIntervals mobj������Ϣ��.�Ƿ���ſ���, Nothing
    RaiseEvent TimeIntervalsChanged(Get������Ϣ��, True)
End Property

Public Property Get ����ʱ��() As Boolean
    If mobj������Ϣ�� Is Nothing Then Set mobj������Ϣ�� = New ������Ϣ��
    ����ʱ�� = mobj������Ϣ��.�Ƿ��ʱ��
End Property

Public Property Let ����ʱ��(ByVal vNewValue As Boolean)
    If mobj������Ϣ�� Is Nothing Then Set mobj������Ϣ�� = New ������Ϣ��
    mobj������Ϣ��.�Ƿ��ʱ�� = vNewValue
    
    m_IsDataChanged = True: RaiseEvent DataIsChanged
    
    ShowTimeIntervals mobj������Ϣ��.�Ƿ���ſ���, Nothing
    RaiseEvent TimeIntervalsChanged(Get������Ϣ��, True)
End Property

Public Property Get ԤԼ����() As Integer
    If mobj������Ϣ�� Is Nothing Then Set mobj������Ϣ�� = New ������Ϣ��
    ԤԼ���� = mobj������Ϣ��.ԤԼ����
End Property

Public Property Let ԤԼ����(ByVal vNewValue As Integer)
    If mobj������Ϣ�� Is Nothing Then Set mobj������Ϣ�� = New ������Ϣ��
    mobj������Ϣ��.ԤԼ���� = vNewValue
    
    cmdFun(3).Enabled = (m_EditMode = ED_RegistPlan_Edit Or m_EditMode = ED_RegistPlan_NumLimitModify) And mobj������Ϣ��.ԤԼ���� <> 1
    cmdFun(4).Enabled = (m_EditMode = ED_RegistPlan_Edit Or m_EditMode = ED_RegistPlan_NumLimitModify) And mobj������Ϣ��.ԤԼ���� <> 1
    vsTimeWork.Editable = IIf((m_EditMode = ED_RegistPlan_Edit Or m_EditMode = ED_RegistPlan_NumLimitModify) And mobj������Ϣ��.ԤԼ���� <> 1, flexEDKbdMouse, flexEDNone)
    
    '���ذ�ť
    cmdԤԼ.Visible = False
End Property

Private Sub SetCtrlMove()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ƶ��ؼ�
    '����:���˺�
    '����:2016-01-13 14:23:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln������� As Boolean, bln����ʱ�� As Boolean
    Dim blnDel As Boolean, lng��Լ As Long
    Dim lngRow As Long
    Dim lngCol As Long
    Err = 0: On Error GoTo Errhand:
    
    bln������� = True: bln����ʱ�� = True
    lng��Լ = 0
    If Not mobj������Ϣ�� Is Nothing Then
        bln������� = mobj������Ϣ��.�Ƿ���ſ���
        bln����ʱ�� = mobj������Ϣ��.�Ƿ��ʱ��
        lng��Լ = mobj������Ϣ��.��Լ��
    End If
    
    cmdԤԼ.Visible = False
    cmdɾ��.Visible = False
    If Not (bln������� And bln����ʱ��) Then Exit Sub
    
    With vsTimeWork
        If .Col < 0 And .Cols > 2 Then .Col = 1
        If .Col < 0 Or .Row < 0 Then Exit Sub
        If .TextMatrix(.Row, .Col) = "" Then Exit Sub
        If .Cell(flexcpFontStrikethru, .Row, .Col) Then Exit Sub  '��ɾ���ߵı�ʾ����Ҫɾ����
        
        lngRow = .Row
        If lngRow Mod 2 = 0 Then lngRow = lngRow + 1
        cmdԤԼ.Left = .CellLeft
        cmdɾ��.Left = .CellLeft + .CellWidth - cmdɾ��.Width - 15
        If .Row Mod 2 = 0 Then
            cmdԤԼ.Top = .CellTop
            cmdɾ��.Top = .CellTop
        Else
            cmdԤԼ.Top = .Cell(flexcpTop, .Row - 1, .Col)
            cmdɾ��.Top = cmdԤԼ.Top
        End If
        
        cmdԤԼ.Visible = lng��Լ <> 0
        cmdԤԼ.Refresh '��ֹ��ť������
        
        'ɾ����ť���һ�в���ʾ
        For lngCol = .Cols - 1 To 1 Step -1
            If Trim(.TextMatrix(.Row, lngCol)) <> "" Then
                If lngCol = .Col Then
                    cmdɾ��.Visible = True
                    cmdɾ��.Refresh '��ֹ��ť������
                End If
                Exit For
            End If
        Next
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub UserControl_Terminate()
    Set mobj������Ϣ�� = Nothing
    Set mobj�ϰ�ʱ�� = Nothing
    Set mcllFixedSN = Nothing
End Sub

Private Sub vsTimeWork_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    m_IsDataChanged = True: RaiseEvent DataIsChanged
    RaiseEvent TimeIntervalsChanged(Get������Ϣ��, False)
End Sub

Private Sub vsTimeWork_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If Not (m_EditMode = ED_RegistPlan_Edit Or m_EditMode = ED_RegistPlan_NumLimitModify) Then Exit Sub
    Call SetCtrlMove
End Sub

Private Sub DeleteTime(ByVal lngRow As Long, ByVal lngCol As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ɾ��ʱ���
    '����:���˺�
    '����:2016-01-13 15:13:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, lngSkip As Long
    Err = 0: On Error GoTo Errhand:
    With vsTimeWork
        If lngRow Mod 2 = 1 Then lngRow = lngRow - 1
        If lngCol < 1 Or lngCol > .Cols - 1 Then Exit Sub
        If lngRow < 0 Or lngRow > .Rows - 1 Then Exit Sub
        If lngCol = .Cols - 1 Then
            .TextMatrix(lngRow, lngCol) = ""
            .TextMatrix(lngRow + 1, lngCol) = ""
            .Cell(flexcpData, lngRow, lngCol) = ""
            .Cell(flexcpData, lngRow + 1, lngCol) = ""
            .Cell(flexcpForeColor, lngRow, lngCol, lngRow + 1, lngCol) = &H80000008
        Else
            For i = lngCol To .Cols - 2
                lngSkip = i + 1
                .TextMatrix(lngRow, i) = .TextMatrix(lngRow, lngSkip)
                .TextMatrix(lngRow + 1, i) = .TextMatrix(lngRow + 1, lngSkip)
                .Cell(flexcpData, lngRow, i) = .Cell(flexcpData, lngRow, lngSkip)
                .Cell(flexcpData, lngRow + 1, i) = .Cell(flexcpData, lngRow + 1, lngSkip)
                .Cell(flexcpForeColor, lngRow, i, lngRow + 1, i) = .Cell(flexcpForeColor, lngRow, lngSkip, lngRow + 1, lngSkip)
                
                .TextMatrix(lngRow, lngSkip) = ""
                .TextMatrix(lngRow + 1, lngSkip) = ""
                .Cell(flexcpData, lngRow, lngSkip) = ""
                .Cell(flexcpData, lngRow + 1, lngSkip) = ""
                .Cell(flexcpForeColor, lngRow, lngSkip, lngRow + 1, lngSkip) = &H80000008
            Next
        End If
    End With
    Call ReSetNumNo
    If vsTimeWork.TextMatrix(lngRow, lngCol) = "" Then cmdɾ��.Visible = False: cmdԤԼ.Visible = False
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub ReSetNumNo()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���µ������
    '����:���˺�
    '����:2016-01-13 15:20:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, lngCol As Long, lngNumNo As Long
    
    With vsTimeWork
        lngNumNo = 0
        For lngRow = 0 To .Rows - 1 Step 2
            For lngCol = 1 To .Cols - 1
               If Trim(.TextMatrix(lngRow, lngCol)) <> "" Then
                    lngNumNo = lngNumNo + 1
                    .TextMatrix(lngRow, lngCol) = lngNumNo
               End If
            Next
        Next
    End With
End Sub
 
'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "����/���ö������ı���ͼ�εı���ɫ��"
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor = New_BackColor
    PropertyChanged "BackColor"
    SetBackColor Controls, UserControl.BackColor
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "ָ�� Label �� Shape �ı�����ʽ��͸���Ļ��ǲ�͸���ġ�"
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "����һ�� Font ����"
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=0,0,0,0
Public Property Get CanReCalic() As Boolean
    CanReCalic = m_CanReCalic
End Property

Public Property Let CanReCalic(ByVal New_CanReCalic As Boolean)
    m_CanReCalic = New_CanReCalic
    PropertyChanged "CanReCalic"
    picFunBack.Visible = m_CanReCalic
    Call UserControl_Resize
End Property

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "����/���õ���꾭������ĳһ����ʱ����ָ�����͡�"
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'Ϊ�û��ؼ���ʼ������
Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    m_CanReCalic = m_def_CanReCalic
    m_EditMode = m_def_EditMode
    m_IsDataChanged = m_def_IsDataChanged
    m_����Ƶ�� = m_def_����Ƶ��
    txtUpd.Text = IIf(m_����Ƶ�� = 0, m_def_����Ƶ��, m_����Ƶ��)
End Sub

'�Ӵ������м�������ֵ
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_CanReCalic = PropBag.ReadProperty("CanReCalic", m_def_CanReCalic)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    m_EditMode = PropBag.ReadProperty("EditMode", m_def_EditMode)
    m_IsDataChanged = PropBag.ReadProperty("IsDataChanged", m_def_IsDataChanged)
    m_����Ƶ�� = PropBag.ReadProperty("����Ƶ��", m_def_����Ƶ��)
    txtUpd.Text = IIf(m_����Ƶ�� = 0, m_def_����Ƶ��, m_����Ƶ��)
End Sub

'������ֵд���洢��
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("CanReCalic", m_CanReCalic, m_def_CanReCalic)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("EditMode", m_EditMode, m_def_EditMode)
    Call PropBag.WriteProperty("IsDataChanged", m_IsDataChanged, m_def_IsDataChanged)
    Call PropBag.WriteProperty("����Ƶ��", m_����Ƶ��, m_def_����Ƶ��)

End Sub

Private Sub vsTimeWork_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    cmdԤԼ.Visible = False
    cmdɾ��.Visible = False
End Sub

Private Sub vsTimeWork_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not (m_EditMode = ED_RegistPlan_Edit Or m_EditMode = ED_RegistPlan_NumLimitModify) Then Cancel = True: Exit Sub
    If mobj������Ϣ�� Is Nothing Then Cancel = True: Exit Sub
    If mobj������Ϣ��.�Ƿ���ſ��� Then Cancel = True: Exit Sub
    If Col Mod 2 = 0 Then Cancel = True: Exit Sub
    If vsTimeWork.Cell(flexcpData, Row, Col) = "" Then Cancel = True: Exit Sub
End Sub

Private Sub vsTimeWork_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And vsTimeWork.Editable = flexEDKbdMouse Then
        If vsTimeWork.Row = vsTimeWork.Rows - 1 And vsTimeWork.Col = vsTimeWork.Cols - 1 Then
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            Call ToNextGridPostion(vsTimeWork, 1, 2, 1, 1)
        End If
        KeyCode = 0
    End If
End Sub

Private Sub ToNextGridPostion(vsfGrid As VSFlexGrid, Optional ByVal lngStepRow As Long = 1, Optional ByVal lngStepCol As Long = 1, _
    Optional ByVal lngFirstRow As Long, Optional ByVal lngFirstCol As Long)
    '���ܣ��Զ�������һ����Ԫ��
    '��Σ�
    '   lngStepRow - �����
    '   lngStepCol - �����
    '   lngFirstRow - ��һ��
    '   lngFirstCol - ��һ��
    Dim lngCurRow As Long, lngCurCol As Long
    With vsfGrid
        If lngFirstRow < vsfGrid.FixedRows Then lngFirstRow = vsfGrid.FixedRows
        If lngFirstCol < vsfGrid.FixedCols Then lngFirstCol = vsfGrid.FixedCols
        
        lngCurRow = .Row: lngCurCol = .Col
        If lngCurCol < lngFirstCol Then
            lngCurCol = lngFirstCol
            .Col = lngCurCol
            Exit Sub
        End If
        
        If (lngCurCol - lngFirstCol) Mod lngStepCol <> 0 Then
            lngCurCol = lngFirstCol + (lngCurCol - lngFirstCol) \ lngStepCol * lngStepCol
            If lngCurCol < lngFirstCol Then lngCurCol = lngFirstCol
        End If
        'ȷ����һ����
        If lngCurCol + lngStepCol > .Cols - 1 Then
            lngCurCol = lngFirstCol
            'ȷ����һ����
            If lngCurRow + lngStepRow > .Rows - 1 Then
                lngCurRow = lngFirstRow
            Else
                lngCurRow = lngCurRow + lngStepRow
            End If
        Else
            lngCurCol = lngCurCol + lngStepCol
        End If
        .Row = lngCurRow: .Col = lngCurCol
    End With
End Sub

Private Sub vsTimeWork_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub vsTimeWork_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack Then Exit Sub
    '����λ�����ƣ�����λ���Ȳ��ܴ���9
    If InStr(vsTimeWork.EditText, ".") > 0 Then
        If InStr(vsTimeWork.EditText, ".") > 9 Then KeyAscii = 0
    Else
        If Len(vsTimeWork.EditText) >= 9 Then KeyAscii = 0
    End If
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Public Function IsValied(Optional ByVal blnChanged As Boolean) As Boolean
    '�������
    '����һ���Ƿ�ı䣬���ı��򱾲�ҲҪ���
    Dim lngSum As Long, lngCount As Long
    Dim bln������� As Boolean, bln����ʱ�� As Boolean
    Dim lng�޺��� As Long, lng��Լ�� As Long
    Dim bytԤԼ���� As Byte, str�ϰ�ʱ�� As String
    Dim lngRow As Long, lngCol As Long, blnԤԼ As Boolean
    Dim strFirstStart As String
    
    Err = 0: On Error GoTo ErrHandler
    '����δ�ı䲻���
    If m_IsDataChanged = False And blnChanged = False Then IsValied = True: Exit Function
    If mobj������Ϣ�� Is Nothing Then IsValied = True: Exit Function
    
    If Not mobj�ϰ�ʱ�� Is Nothing Then str�ϰ�ʱ�� = mobj�ϰ�ʱ��.ʱ���
    bln������� = mobj������Ϣ��.�Ƿ���ſ���
    bln����ʱ�� = mobj������Ϣ��.�Ƿ��ʱ��
    lng�޺��� = mobj������Ϣ��.�޺���
    lng��Լ�� = IIf(mobj������Ϣ��.ԤԼ���� = 1, 0, _
            IIf(mobj������Ϣ��.��Լ�� = 0 And mobj������Ϣ��.�޺��� <> 0, mobj������Ϣ��.�޺���, mobj������Ϣ��.��Լ��))
    bytԤԼ���� = mobj������Ϣ��.ԤԼ����
    
    If bln����ʱ�� = False Then IsValied = True: Exit Function
    '----------------------------------------------------------------
    '���⴦�����������ڱ༭״̬ʱ����鲻��
    mblnValiedCanSave = True
    vsTimeWork.FinishEditing False
    If mblnValiedCanSave = False Then
        Exit Function
    Else
        mblnValiedCanSave = False
    End If
    '----------------------------------------------------------------
    
    With vsTimeWork
        If Not bln������� Then
            For lngRow = 1 To .Rows - 1
                For lngCol = 0 To .Cols - 1 Step 2
                    If .TextMatrix(lngRow, lngCol) <> "" Then
                        lngCount = lngCount + 1
                        lngSum = lngSum + Val(.TextMatrix(lngRow, lngCol + 1))
                        
                        If lngRow = 1 And lngCol = 0 Then
                            strFirstStart = Split(.Cell(flexcpData, lngRow, lngCol + 1), "��")(0)
                        End If
                    End If
                Next
            Next
            If lngSum = 0 And bytԤԼ���� <> 1 Then
                MsgBox IIf(m_EditMode = ED_RegistPlan_NumLimitModify, "", str�ϰ�ʱ��) & "������ʱ�������Ҫ������Լʱ�Σ�", vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
            If lngSum > lng��Լ�� Then
                MsgBox IIf(m_EditMode = ED_RegistPlan_NumLimitModify, "", str�ϰ�ʱ�� & "��") & "��ԤԼ����(" & lngSum & ")��������Լ��(" & lng��Լ�� & ")��", vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
        Else
            For lngRow = 0 To .Rows - 1 Step 2
                For lngCol = 1 To .Cols - 1
                    If .TextMatrix(lngRow, lngCol) <> "" Then
                        lngCount = lngCount + 1
                        blnԤԼ = Val(.Cell(flexcpData, lngRow, lngCol)) = 1
                        If blnԤԼ Then lngSum = lngSum + 1
                        
                        If lngRow = 0 And lngCol = 1 Then
                            strFirstStart = Split(.Cell(flexcpData, lngRow + 1, lngCol), "��")(0)
                        End If
                    End If
                Next
            Next
            'If lngSum = 0 Then lngSum = lngCount '�������ʾȫ��ԤԼ
            If lngSum = 0 And bytԤԼ���� <> 1 Then
                MsgBox IIf(m_EditMode = ED_RegistPlan_NumLimitModify, "", str�ϰ�ʱ��) & "������ʱ�������Ҫ���ÿ�ԤԼʱ�Σ�", vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
            If lngSum < lng��Լ�� Then
                If MsgBox("ע�⣺" & vbCrLf & _
                    "        " & IIf(m_EditMode = ED_RegistPlan_NumLimitModify, "", str�ϰ�ʱ��) & " ��ԤԼʱ��ε�����(" & lngSum & ")����Լ��(" & lng��Լ�� & ")���ȣ���ȷ������ǰ���ñ�����", _
                    vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            End If
        End If
    End With
    
    If lngCount = 0 Then
        MsgBox "����ʱ��ʱ����Ҫ����ʱ�Σ�", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If Not mobj�ϰ�ʱ�� Is Nothing Then
        '����һ��ʱ�εĿ�ʼʱ���Ƿ�����ϰ�ʱ�εĿ�ʼʱ��
        If IsDate(strFirstStart) Then
            strFirstStart = Format(mobj�ϰ�ʱ��.��ʼʱ��, "yyyy-mm-dd") & " " & Format(strFirstStart, "hh:mm:ss")
            If DateDiff("n", mobj�ϰ�ʱ��.��ʼʱ��, strFirstStart) <> 0 Then
                If MsgBox(mobj�ϰ�ʱ��.ʱ��� & " ��һ�����ʱ�εĿ�ʼʱ��(" & Format(strFirstStart, "hh:mm") & _
                    ")�뵱ǰ�ϰ�ʱ�εĿ�ʼʱ��(" & Format(mobj�ϰ�ʱ��.��ʼʱ��, "hh:mm") & _
                    ")��ͬ����ȷ������ǰ���ñ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            End If
        End If
    End If
    IsValied = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub vsTimeWork_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lngSum As Long
    Dim lngRow As Long, lngCol As Long
    Dim lng��ԤԼ As Long
    
    On Error GoTo Errhand
    If mobj������Ϣ�� Is Nothing Then Cancel = True: Exit Sub
    If mobj������Ϣ��.�Ƿ���ſ��� Then Cancel = True: Exit Sub
    
    With vsTimeWork
        '����λ����9λ��ֱ�ӽص�,��ֹ���
        If InStr(.EditText, ".") > 0 Then
            If InStr(.EditText, ".") > 9 Then
                 .EditText = Left(.EditText, 9)
            End If
        Else
             .EditText = Left(.EditText, 9)
        End If
    
        If ValiedCanModify(Val(.Cell(flexcpData, .Row, .Col - 1)), Val(.EditText), False, lng��ԤԼ) = False Then
            MsgBox "��ǰʱ����ԤԼ " & lng��ԤԼ & " ������ԤԼ������С�� " & lng��ԤԼ & " ��", vbInformation, gstrSysName
            Cancel = True: mblnValiedCanSave = False: Exit Sub
        End If
        
        For lngRow = 1 To .Rows - 1
            For lngCol = 0 To .Cols - 1 Step 2
                If .TextMatrix(lngRow, lngCol) <> "" Then
                    If lngRow = Row And lngCol = Col - 1 Then
                        lngSum = lngSum + Val(.EditText)
                    Else
                        lngSum = lngSum + Val(.TextMatrix(lngRow, lngCol + 1))
                    End If
                End If
            Next
        Next
        If lngSum > mobj������Ϣ��.��Լ�� Then
            If Val(.EditText) < Val(.TextMatrix(.Row, .Col)) Or Val(.EditText) = 0 Then
                Exit Sub
            Else
                MsgBox "ԤԼ��(" & lngSum & ")���ܳ�����Լ��(" & mobj������Ϣ��.��Լ�� & ")��", vbInformation + vbOKOnly, gstrSysName
                Cancel = True: mblnValiedCanSave = False: Exit Sub
            End If
        End If
        .EditText = FormatEx(Val(.EditText), 0)
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub SetNewSN(ByVal lng�޺��� As Long, ByVal lngCurAdd As Long, ByVal blnAdd As Boolean)
    '�ӺŻ��߼���
    '����
    '   lng�޺�������ǰ�޺���
    '   lngCurAdd�����������޺������Ӻ�Ϊ��������Ϊ��
    '   blnAdd���Ƿ�ӺŲ���
    Dim dtStart As Date, dtEnd As Date, intStep As Integer
    Dim colNumAll As Collection, objColAll As New Collection
    Dim obj������Ϣ As ������Ϣ
    Dim lngRow As Long, lngCol As Long, lngCount As Long
    Dim dtOriginalStartTime As Date
    
    Err = 0: On Error GoTo ErrHandler
    m_IsDataChanged = True
    mobj������Ϣ��.�޺��� = lng�޺���
    If mblnClickedFunBtn Then
        Call cmdFun_Click(mintPreSelFun)
        Exit Sub
    End If
    If blnAdd Then  '�Ӻ�
        If mobj������Ϣ��.Count > 0 And mobj������Ϣ��.�Ƿ��ʱ�� And mobj������Ϣ��.�Ƿ���ſ��� Then

            intStep = DateDiff("n", mobj������Ϣ��(mobj������Ϣ��.Count).��ʼʱ��, mobj������Ϣ��(mobj������Ϣ��.Count).��ֹʱ��)
            dtStart = mobj������Ϣ��(mobj������Ϣ��.Count).��ֹʱ��
            dtOriginalStartTime = Format(dtStart, "yyyy-mm-dd ") & Format(mobj�ϰ�ʱ��.��ʼʱ��, "hh:mm:ss")
            dtEnd = Format(dtStart, "yyyy-mm-dd ") & Format(mobj�ϰ�ʱ��.����ʱ��, "hh:mm:ss")
            If DateDiff("n", dtEnd, dtStart) > 0 Then dtEnd = DateAdd("d", 1, dtEnd)
            
            '��ȥԤ��ʱ��
            Call ��ȥԤ��ʱ��(dtStart, dtEnd, mobj�ϰ�ʱ��.����Ԥ��ʱ��, mobj�ϰ�ʱ��.��Ϣʱ��)
            If DateDiff("n", dtStart, dtEnd) > 0 Then
                For Each obj������Ϣ In mobj������Ϣ��
                    objColAll.Add Array(obj������Ϣ.���, obj������Ϣ.��ʼʱ��, obj������Ϣ.��ֹʱ��, IIf(obj������Ϣ.�Ƿ�ԤԼ, 1, 0))
                Next
                Set colNumAll = CalculatTimeInterval(0, mobj������Ϣ��.�Ƿ���ſ���, intStep, objColAll.Count + lngCurAdd, _
                    dtStart, dtEnd, mobj�ϰ�ʱ��.��Ϣʱ��, objColAll.Count + 1, , Format(dtOriginalStartTime, "yyyy-MM-dd hh:mm:ss"))
                AddRange objColAll, colNumAll
                ShowTimeIntervals mobj������Ϣ��.�Ƿ���ſ���, objColAll
                
                lngCount = mobj������Ϣ��.Count
                For lngRow = 0 To vsTimeWork.Rows - 1 Step 2
                    For lngCol = 1 To vsTimeWork.Cols - 1 Step 1


                        If vsTimeWork.TextMatrix(lngRow, lngCol) <> "" Then
                            If lngCount <= 0 Then
                                vsTimeWork.Cell(flexcpForeColor, lngRow, lngCol, lngRow + 1, lngCol) = vbMagenta
                            End If
                            lngCount = lngCount - 1
                        End If
                    Next
                Next
            End If
        End If
    Else  '����
        If mobj������Ϣ��.Count > 0 And mobj������Ϣ��.�Ƿ��ʱ�� And mobj������Ϣ��.�Ƿ���ſ��� Then
            For Each obj������Ϣ In mobj������Ϣ��
                objColAll.Add Array(obj������Ϣ.���, obj������Ϣ.��ʼʱ��, obj������Ϣ.��ֹʱ��, IIf(obj������Ϣ.�Ƿ�ԤԼ, 1, 0))
            Next
            ShowTimeIntervals mobj������Ϣ��.�Ƿ���ſ���, objColAll
            
            lngCount = lng�޺���
            For lngRow = 0 To vsTimeWork.Rows - 1 Step 2
                For lngCol = 1 To vsTimeWork.Cols - 1 Step 1


                    If vsTimeWork.TextMatrix(lngRow, lngCol) <> "" Then
                        If lngCount <= 0 Then
                            If ValiedCanModify(Val(vsTimeWork.TextMatrix(lngRow, lngCol)), 0, True) Then
                                vsTimeWork.Cell(flexcpForeColor, lngRow, lngCol, lngRow + 1, lngCol) = vbRed
                                vsTimeWork.Cell(flexcpFontStrikethru, lngRow, lngCol, lngRow + 1, lngCol) = True
                            End If
                        End If
                        lngCount = lngCount - 1
                    End If
                Next
            Next
        End If
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function ��ȥԤ��ʱ��(ByRef dtStartDate As Date, ByRef dtEndDate As Date, _
    ByVal lngԤ��ʱ�� As Long, ByVal str��Ϣʱ�� As String) As Boolean
    '��ȥԤ��ʱ��
    Dim dtStart As Date, dtEnd As Date
    Dim i As Integer, varTemp As Variant
    
    Err = 0: On Error GoTo ErrHandler
    dtEndDate = DateAdd("n", -1 * lngԤ��ʱ��, dtEndDate)
    varTemp = Split(str��Ϣʱ��, ";")
    For i = 0 To UBound(varTemp)
        '�����Ϣʱ�εĿ�ʼʱ��С���ϰ�ʱ�εĿ�ʼʱ�䣬���ʾ�ǵڶ��죬��Ϣʱ�εĿ�ʼʱ�����ֹʱ�䶼Ҫ��һ��
        dtStart = CDate(Format(dtStartDate, "yyyy-mm-dd ") & Split(varTemp(i), "-")(0))
        dtEnd = CDate(Format(dtStartDate, "yyyy-mm-dd ") & Split(varTemp(i), "-")(1))
        If DateDiff("n", dtStart, dtStartDate) > 0 Then dtStart = DateAdd("d", 1, dtStart): dtEnd = DateAdd("d", 1, dtEnd)
        '��Ϣʱ�ε���ֹʱ��С����Ϣʱ�εĿ�ʼʱ�䣬����Ϣʱ�ε���ֹʱ���һ��
        If DateDiff("n", dtEnd, dtStart) > 0 Then dtEnd = DateAdd("d", 1, dtEnd)
        '����ϰ�ʱ�ε���ֹʱ������Ϣʱ���ڣ����ϰ�ʱ�ε���ֹʱ��ȡ��Ϣʱ�εĿ�ʼʱ��
        If DateDiff("n", dtEndDate, dtEnd) <= 0 And DateDiff("n", dtEndDate, dtStart) >= 0 Then dtEndDate = dtStart
    Next
    ��ȥԤ��ʱ�� = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ValiedCanModify(ByVal lng��� As Long, ByVal lng���� As Long, _
    Optional ByVal blnDel As Boolean, Optional ByRef lng��ԤԼ As Long) As Boolean
    '��鵱ǰʱ���Ƿ�����޸�
    Dim i As Long, arrSN As Variant
    
    Err = 0: On Error GoTo ErrHandler
    If mcllFixedSN Is Nothing Then ValiedCanModify = True: Exit Function
    If mcllFixedSN.Count = 0 Then ValiedCanModify = True: Exit Function
    
    For i = 1 To mcllFixedSN.Count
        arrSN = mcllFixedSN(i) '(���,����)
        If blnDel Then
            'ɾ����Ҫ������ʱ�Σ���������ſ���
            If arrSN(0) >= lng��� And lng���� < arrSN(1) Then
                lng��ԤԼ = arrSN(1)
                Exit Function
            End If
        ElseIf arrSN(0) = lng��� And lng���� < arrSN(1) Then
            lng��ԤԼ = arrSN(1)
            Exit Function
        End If
    Next
    ValiedCanModify = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=1,0,0,0
Public Property Get EditMode() As gRegistPlanEditMode
    EditMode = m_EditMode
End Property

Public Property Let EditMode(ByVal New_EditMode As gRegistPlanEditMode)
    m_EditMode = New_EditMode
    PropertyChanged "EditMode"
    
    picFunBack.Visible = m_CanReCalic And (EditMode = ED_RegistPlan_Edit Or m_EditMode = ED_RegistPlan_NumLimitModify)
    Call UserControl_Resize
    SetEnabled UserControl.Controls, m_EditMode = ED_RegistPlan_Edit Or m_EditMode = ED_RegistPlan_NumLimitModify
    SetEnabledBackColor UserControl.Controls
    
    vsTimeWork.Editable = flexEDNone
    If mobj������Ϣ�� Is Nothing Then Exit Property
    SetFunVisible mobj������Ϣ��.�Ƿ���ſ���
    If mobj������Ϣ��.�Ƿ���ſ��� = False And mobj������Ϣ��.�Ƿ��ʱ�� Then
        vsTimeWork.Editable = IIf(m_EditMode = ED_RegistPlan_Edit Or m_EditMode = ED_RegistPlan_NumLimitModify, flexEDKbdMouse, flexEDNone)
        vsTimeWork.Editable = IIf(vsTimeWork.Editable = flexEDKbdMouse And mobj������Ϣ��.ԤԼ���� <> 1, flexEDKbdMouse, flexEDNone)
    End If
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=0,0,0,false
Public Property Get IsDataChanged() As Boolean
    IsDataChanged = m_IsDataChanged
End Property

Public Property Let IsDataChanged(ByVal New_IsDataChanged As Boolean)
    m_IsDataChanged = New_IsDataChanged
    PropertyChanged "IsDataChanged"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=7,0,0,5
Public Property Get ����Ƶ��() As Integer
    ����Ƶ�� = m_����Ƶ��
End Property

Public Property Let ����Ƶ��(ByVal New_����Ƶ�� As Integer)
    m_����Ƶ�� = New_����Ƶ��
    PropertyChanged "����Ƶ��"
    txtUpd.Text = IIf(m_����Ƶ�� = 0, m_def_����Ƶ��, m_����Ƶ��)
End Property

