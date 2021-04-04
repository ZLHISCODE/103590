VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Begin VB.UserControl CalendarSel 
   ClientHeight    =   6210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10560
   ScaleHeight     =   6210
   ScaleWidth      =   10560
   Begin VB.PictureBox picWeekList 
      BorderStyle     =   0  'None
      Height          =   4065
      Left            =   4380
      ScaleHeight     =   4065
      ScaleWidth      =   4665
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   60
      Visible         =   0   'False
      Width           =   4665
      Begin zl9RegEvent.DayCalendar btnDay 
         Height          =   1935
         Left            =   2460
         TabIndex        =   19
         ToolTipText     =   "��סCtrl����һ��ѡ����δ���ð��ŵ�����"
         Top             =   2040
         Visible         =   0   'False
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   3413
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlignmentY      =   1
         MultiSelectionMode=   -1  'True
         BorderStyle     =   0   'False
      End
      Begin zl9RegEvent.WeekCalendar btnWeek 
         Height          =   1365
         Left            =   60
         TabIndex        =   17
         Top             =   2280
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   2408
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlignmentX      =   1
         AlignmentY      =   1
         BorderStyle     =   0   'False
      End
      Begin VB.CheckBox chkValied 
         Caption         =   "����������"
         Height          =   285
         Index           =   0
         Left            =   75
         TabIndex        =   7
         Top             =   690
         Width           =   1305
      End
      Begin VB.CheckBox chkValied 
         Caption         =   "���ղ�����"
         Height          =   285
         Index           =   1
         Left            =   1485
         TabIndex        =   8
         Top             =   690
         Width           =   1305
      End
      Begin VB.PictureBox picLoop 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   60
         ScaleHeight     =   975
         ScaleWidth      =   3825
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1050
         Width           =   3825
         Begin VB.OptionButton optLoop 
            Caption         =   "������"
            Height          =   225
            Index           =   0
            Left            =   0
            TabIndex        =   9
            Top             =   -15
            Width           =   870
         End
         Begin VB.OptionButton optLoop 
            Caption         =   "������ѭ"
            Height          =   225
            Index           =   1
            Left            =   1395
            TabIndex        =   10
            Top             =   -15
            Value           =   -1  'True
            Width           =   1110
         End
         Begin VB.TextBox txtSkip 
            Height          =   285
            Left            =   405
            TabIndex        =   15
            Text            =   "3"
            Top             =   630
            Width           =   390
         End
         Begin MSComCtl2.UpDown updSkip 
            Height          =   285
            Left            =   765
            TabIndex        =   16
            Top             =   630
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txtSkip"
            BuddyDispid     =   196615
            OrigLeft        =   750
            OrigTop         =   630
            OrigRight       =   1005
            OrigBottom      =   915
            Max             =   100
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.DTPicker dtpLoopDay 
            Height          =   300
            Left            =   1185
            TabIndex        =   13
            Top             =   255
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   158728195
            CurrentDate     =   42377
         End
         Begin VB.ComboBox cboDays 
            Height          =   300
            Left            =   1185
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   255
            Width           =   1260
         End
         Begin VB.Label lblLoopDate 
            AutoSize        =   -1  'True
            Caption         =   "��ʼ��ѭ����"
            Height          =   180
            Left            =   0
            TabIndex        =   11
            Top             =   315
            Width           =   1080
         End
         Begin VB.Label lblLoopSkipDays 
            AutoSize        =   -1  'True
            Caption         =   "���        ��"
            Height          =   180
            Left            =   0
            TabIndex        =   14
            Top             =   675
            Width           =   1260
         End
      End
      Begin VB.OptionButton optRule 
         Caption         =   "����"
         Height          =   240
         Index           =   1
         Left            =   900
         TabIndex        =   3
         Top             =   0
         Width           =   660
      End
      Begin VB.OptionButton optRule 
         Caption         =   "����"
         Height          =   240
         Index           =   0
         Left            =   60
         TabIndex        =   2
         Top             =   0
         Value           =   -1  'True
         Width           =   810
      End
      Begin VB.OptionButton optRule 
         Caption         =   "�ض�����"
         Height          =   240
         Index           =   4
         Left            =   900
         TabIndex        =   6
         Top             =   270
         Width           =   1050
      End
      Begin VB.OptionButton optRule 
         Caption         =   "˫��"
         Height          =   240
         Index           =   2
         Left            =   1890
         TabIndex        =   4
         Top             =   0
         Width           =   660
      End
      Begin VB.OptionButton optRule 
         Caption         =   "��ѭ"
         Height          =   240
         Index           =   3
         Left            =   60
         TabIndex        =   5
         Top             =   270
         Width           =   810
      End
      Begin VB.Line lnDown 
         BorderColor     =   &H80000005&
         X1              =   0
         X2              =   4425
         Y1              =   585
         Y2              =   585
      End
      Begin VB.Line lnTop 
         BorderColor     =   &H80000000&
         X1              =   -210
         X2              =   4605
         Y1              =   570
         Y2              =   570
      End
   End
   Begin XtremeCalendarControl.DatePicker dtpDays 
      Height          =   3030
      Left            =   15
      TabIndex        =   0
      Top             =   60
      Width           =   4245
      _Version        =   1048579
      _ExtentX        =   7488
      _ExtentY        =   5345
      _StockProps     =   64
      AutoSize        =   0   'False
      ShowTodayButton =   0   'False
      ShowNoneButton  =   0   'False
      HighlightToday  =   0   'False
      ShowNonMonthDays=   0   'False
      Show3DBorder    =   0
      MaxSelectionCount=   1
      AskDayMetrics   =   -1  'True
   End
End
Attribute VB_Name = "CalendarSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum gShowStyle
    Show_Plan_Day = 0
    Show_Plan_Week = 1
    Show_Plan_Rule = 2
End Enum

'ȱʡ����ֵ:
Const m_def_BackColor = 0
Const m_def_ForeColor = 0
Const m_def_Enabled = 0
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0
Const m_def_ShowStyle = 0
'���Ա���:
Dim m_BackColor As Long
Dim m_ForeColor As Long
Dim m_Enabled As Boolean
Dim m_Font As Font
Dim m_BackStyle As Integer
Dim m_BorderStyle As Integer
Dim m_ShowStyle As gShowStyle
'�¼�����:
Public Event SelectedChangeBefore(ByVal OldDate As String, NewDate As String, Cancel As Boolean)
Public Event SelectedChanged(ByVal OldDate As String, NewDate As String)

Private Const mDateBefore As String = "2016-01-"
Private mdtMinDay As Date
Private mobj���ﰲ�� As ���ﰲ��

Private mblnNotClick As Boolean
Private mbytKeyShift As Byte '1-vbShiftMask 2-VbCtrlMask 4-vbAltMask
Private mcolPicture As Collection 'ͼƬ���ϣ���ǽڼ��պ�ͣ�ﰲ�ŵġ�
                                '���Ϲؼ���(Valueȡ���ڣ���ʽ"yyyymmdd")��
                                '   K+Value - δѡ�����δ���ð��ŵ�, KS+Value - ��ǰѡ�����δ���ð��ŵ�
                                '   B+Value - δѡ����������˰��ŵ�, BS+Value - ��ǰѡ����������˰��ŵ�

Public Function LoadData(ByRef obj���ﰲ�� As ���ﰲ��) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���س��ﰲ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    Set mobj���ﰲ�� = obj���ﰲ��
    If mobj���ﰲ�� Is Nothing Then Set mobj���ﰲ�� = New ���ﰲ��
    
    LoadData = InitData
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub SetBoldItem(ByVal blnWeek As Boolean)
    Dim i As Integer
    
    If blnWeek Then
        For i = 0 To btnWeek.Count - 1
            btnWeek.ItemBold(i) = False
        Next
    Else
        For i = 0 To btnDay.Count - 1
            btnDay.ItemBold(i) = False
        Next
    End If
    
    If Not mobj���ﰲ��.�ѱ�����ﰲ�� Is Nothing Then
        If mobj���ﰲ��.�ѱ�����ﰲ��.�Ű���� = IIf(blnWeek, 1, 6) Or m_ShowStyle = Show_Plan_Week Then
            If mobj���ﰲ��.�ѱ�����ﰲ��.�Ű���� = mobj���ﰲ��.�Ű���� Then
                For i = 1 To mobj���ﰲ��.�ѱ�����ﰲ��.Count
                    If blnWeek Then
                        If mobj���ﰲ��.�ѱ�����ﰲ��(i).�Ƿ�ɾ�� = False Then
                            btnWeek.ItemBold(GetWeekIndex(mobj���ﰲ��.�ѱ�����ﰲ��(i).��������)) = True
                        End If
                    Else
                        btnDay.ItemBold(Val(mobj���ﰲ��.�ѱ�����ﰲ��(i).��������) - 1) = True
                    End If
                Next
            End If
        End If
    End If
    If Not mobj���ﰲ��.δ������ﰲ�� Is Nothing Then
        For i = 1 To mobj���ﰲ��.δ������ﰲ��.Count
            If blnWeek Then
                btnWeek.ItemBold(GetWeekIndex(mobj���ﰲ��.δ������ﰲ��(i).��������)) = True
            Else
                btnDay.ItemBold(Val(mobj���ﰲ��.δ������ﰲ��(i).��������) - 1) = True
            End If
        Next
    End If
End Sub

Private Function InitData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:���˺�
    '����:2016-01-12 15:36:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim obj�����¼�� As �����¼��
    
    Err = 0: On Error GoTo Errhand:
    
    If m_ShowStyle = Show_Plan_Day Then
        Call AdjustWeekSeat
        dtpDays.SetRange Format(mobj���ﰲ��.��ʼʱ��, "yyyy-mm-dd"), _
            Format(mobj���ﰲ��.��ֹʱ��, "yyyy-mm-dd")
        mdtMinDay = Format(mobj���ﰲ��.��ʼʱ��, "yyyy-mm-dd")
    End If
    
    btnWeek.Tag = ""
    If mobj���ﰲ��.���º�����λ Then
        Enabled = False '�������������
        If m_ShowStyle = Show_Plan_Rule Then
            '�������ѡ
            btnDay.MultiSelectionMode = False
            btnDay.ToolTipText = ""
        End If
    End If
    
    mblnNotClick = True
    dtpDays.ClearSelection
    dtpDays.RedrawControl
    
    optRule(0).Value = True
    optLoop(1).Value = True
    cboDays.ListIndex = 0
    dtpLoopDay.Value = Now
    txtSkip.Text = "1"
    
    btnWeek.ClearAll
    btnDay.ClearAll
    
    If mobj���ﰲ��.Count = 0 Then
        '���������¼��
         Set obj�����¼�� = New �����¼��
        '����ȱʡֵ
        If m_ShowStyle = Show_Plan_Rule And mobj���ﰲ��.�Ű���� <> 1 Then
            mobj���ﰲ��.�Ű���� = 1
            mobj���ﰲ��.ȱʡ�������� = "��һ"
        End If
        With obj�����¼��
            .�������� = mobj���ﰲ��.ȱʡ��������
        End With
        mobj���ﰲ��.AddItem obj�����¼��, GetPlanKey(obj�����¼��.��������)
    End If
    
    '��������
    Select Case m_ShowStyle
    Case Show_Plan_Rule
        '1-�����Ű�;2-�����Ű�;3-˫���Ű�;4-������ѭ;5-��ѭ������;6-�ض�����
        Select Case mobj���ﰲ��.�Ű����
        Case 1
            Call SetBoldItem(True)
            If mobj���ﰲ��(1).�������� = "" Then
                btnWeek.ItemValue(0) = True
                mobj���ﰲ��(1).�������� = GetWeekName(0)
            Else
                btnWeek.ItemValue(GetWeekIndex(mobj���ﰲ��(1).��������)) = True
            End If
        Case 2, 3
            optRule(IIf(mobj���ﰲ��.�Ű���� = 2, 1, 2)).Value = True
            '����������
            If mobj���ﰲ��.���������� Then chkValied(0).Value = vbChecked
            '���ղ�����
            If mobj���ﰲ��.���ղ����� Then chkValied(1).Value = vbChecked
        Case 4, 5
            optRule(3).Value = True
            If mobj���ﰲ��.�Ű���� = 4 Then
                optLoop(1).Value = True
                zlControl.CboLocate cboDays, Val(Format(mobj���ﰲ��.��ʼʱ��, "dd")), True
                txtSkip.Text = Val(mobj���ﰲ��(1).��������)
            Else
                optLoop(0).Value = True
                dtpLoopDay.Value = Format(mobj���ﰲ��.��ʼʱ��, "yyyy-mm-dd")
            End If
            txtSkip.Text = Val(mobj���ﰲ��(1).��������)
            '����������
            If mobj���ﰲ��.���������� Then chkValied(0).Value = vbChecked
            '���ղ�����
            If mobj���ﰲ��.���ղ����� Then chkValied(1).Value = vbChecked
            
            dtpLoopDay.Visible = optLoop(0).Value And m_ShowStyle = Show_Plan_Rule
            cboDays.Visible = Not optLoop(0).Value And m_ShowStyle = Show_Plan_Rule
        Case 6
            optRule(4).Value = True
            Call SetBoldItem(False)
            If mobj���ﰲ��(1).�������� = "" Then
                btnDay.ItemValue(0) = True
                mobj���ﰲ��(1).�������� = btnDay.ItemCaption(0) & "��"
            Else
                btnDay.ItemValue(Val(mobj���ﰲ��(1).��������) - 1) = True
            End If
        End Select
    Case Show_Plan_Week
        Call SetBoldItem(True)
        If mobj���ﰲ��(1).�������� = "" Then
            btnWeek.ItemValue(0) = True
            mobj���ﰲ��(1).�������� = GetWeekName(0)
        Else
            btnWeek.ItemValue(GetWeekIndex(mobj���ﰲ��(1).��������)) = True
        End If
    Case Show_Plan_Day
        If mobj���ﰲ��(1).�������� = "" Then
            mobj���ﰲ��(1).�������� = mobj���ﰲ��.��ʼʱ��
            dtpDays.Select CDate(mobj���ﰲ��(1).��������)
            dtpDays.EnsureVisibleSelection
            dtpDays.RedrawControl
            Call dtpDays_SelectionChanged
        ElseIf IsDate(mobj���ﰲ��(1).��������) Then
            dtpDays.Select CDate(mobj���ﰲ��(1).��������)
            dtpDays.EnsureVisibleSelection
            dtpDays.RedrawControl
            Call dtpDays_SelectionChanged
        End If
    End Select
    mblnNotClick = False
    InitData = True
    Exit Function
Errhand:
    mblnNotClick = False
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub AdjustWeekSeat()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ڵ�λ��
    '����:���˺�
    '����:2016-01-08 15:18:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, sngTop As Single, sngLeft As Single
    Dim sngRuleTop As Single, sngRuleLeft As Single

    Err = 0: On Error GoTo Errhand:
    sngRuleTop = 0: sngRuleLeft = 0
    Select Case m_ShowStyle
    Case Show_Plan_Rule '����
        picWeekList.Visible = True
        For i = 0 To optLoop.UBound
            optLoop(i).Visible = True
        Next
        For i = 0 To optRule.UBound
            optRule(i).Visible = True
        Next
        sngRuleTop = optRule(4).Top + optRule(4).Height + 100
        sngRuleLeft = optRule(0).Left
        btnWeek.Visible = optRule(0).Value
        chkValied(0).Visible = Not optRule(0).Value And Not optRule(4).Value
        chkValied(1).Visible = Not optRule(0).Value And Not optRule(4).Value
        picLoop.Visible = optRule(3).Value
        dtpLoopDay.Visible = optLoop(0).Value
        picLoop.Top = chkValied(1).Top + chkValied(1).Height + 50
        cboDays.Visible = Not optLoop(0).Value
        lnTop.Visible = True
        lnDown.Visible = True
        dtpDays.Visible = False
        btnDay.Visible = optRule(4).Value
    Case Show_Plan_Day  '������
        btnWeek.Visible = False
        picWeekList.Visible = False
        Set dtpDays.Container = Me '���ڿؼ���
        dtpDays.Visible = True
        dtpDays.Left = ScaleLeft
        dtpDays.Width = ScaleWidth
        If Not mobj���ﰲ�� Is Nothing Then
            If mobj���ﰲ��.ģ������ = 2 Then
                dtpDays.Top = -550
            Else
                dtpDays.Top = ScaleTop
            End If
        Else
            dtpDays.Top = ScaleTop
        End If
        dtpDays.Height = ScaleHeight - dtpDays.Top
    Case Show_Plan_Week  '������
        For i = 0 To optLoop.UBound
            optLoop(i).Visible = False
        Next
        For i = 0 To optRule.UBound
            optRule(i).Visible = False
        Next
        picWeekList.Visible = True
        btnWeek.Visible = True
        dtpDays.Visible = False
        chkValied(0).Visible = False
        chkValied(1).Visible = False
        picLoop.Visible = False
        dtpLoopDay.Visible = False
        lnTop.Visible = False
        lnDown.Visible = False
    End Select
    picLoop.Top = chkValied(0).Top + chkValied(0).Height + 50
    picLoop.Left = chkValied(0).Left
    Call picWeekList_Resize
    Call UserControl_Resize
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
    '����:2016-01-08 14:44:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    On Error GoTo Errhand
    dtpDays.HighlightToday = False
    dtpDays.ShowNonMonthDays = False
    With dtpDays.PaintManager
        .ControlBackColor = &HFFEFE3
        .DayBackColor = &HFFEFE3
        .DaysOfWeekBackColor = &HFFEFE3
    End With
    
    With cboDays
        .Clear
        For i = 1 To 30
            .AddItem "��" & i & "��"
            .ItemData(.NewIndex) = i
        Next
        mblnNotClick = True
        .ListIndex = 0
        mblnNotClick = False
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub btnWeek_Click(idxWeek As Integer, Value As Boolean)
    Dim i As Integer
    Dim blnCancel As Boolean
    Dim strOldDate As String, strNewDate As String
    
    On Error GoTo Errhand
    If btnWeek.Tag = "" Then btnWeek.Tag = idxWeek
    If mblnNotClick Then Exit Sub
    
    '����ѡ��һ��
    If Value = False Then
        btnWeek.ItemValue(idxWeek) = True
    End If
    
    strOldDate = GetWeekName(Val(btnWeek.Tag)): strNewDate = GetWeekName(idxWeek)
    If btnWeek.Tag = "" Or Val(btnWeek.Tag) <> idxWeek Then
        RaiseEvent SelectedChangeBefore(strOldDate, strNewDate, blnCancel)
        If blnCancel Then
            mblnNotClick = True
            btnWeek.ItemValue(idxWeek) = False
            btnWeek.ItemValue(Val(btnWeek.Tag)) = True
            mblnNotClick = False
            Exit Sub
        End If
        btnWeek.Tag = idxWeek
        Call ChangeCurPlan(mobj���ﰲ��, GetWeekName(idxWeek))
        RaiseEvent SelectedChanged(strOldDate, strNewDate)
        Call SetBoldItem(True)
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub chkValied_Click(index As Integer)
    On Error GoTo Errhand
    If mblnNotClick Then Exit Sub
    If index = 0 Then '����������
        mobj���ﰲ��.���������� = (chkValied(index).Value = vbChecked)
    Else '���ղ�����
        mobj���ﰲ��.���ղ����� = (chkValied(index).Value = vbChecked)
    End If
    '�ֶ�����Ϊ���޸ģ���Ϊ��Ȼ����ı��ˣ����Ǽ�¼�����ܻ���һ��
    If mobj���ﰲ��.Count > 0 Then mobj���ﰲ��(1).�Ƿ��޸� = True
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cboDays_Click()
    Dim blnCancel As Boolean
    
    On Error GoTo Errhand
    If mblnNotClick Then Exit Sub
    mobj���ﰲ��.��ʼʱ�� = CDate(mDateBefore & cboDays.ItemData(cboDays.ListIndex))
    '�ֶ�����Ϊ���޸ģ���Ϊ��Ȼ����ı��ˣ����Ǽ�¼�����ܻ���һ��
    If mobj���ﰲ��.Count > 0 Then mobj���ﰲ��(1).�Ƿ��޸� = True
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub dtpDays_DayMetrics(ByVal Day As Date, ByVal Metrics As XtremeCalendarControl.IDatePickerDayMetrics)
    Dim blnSelected As Boolean, strKey As String
    Dim strStopTxt As String, blnNotSelect As Boolean
    
    On Error GoTo Errhand
    If mobj���ﰲ�� Is Nothing Then Exit Sub
    If mobj���ﰲ��.Count = 0 Then Exit Sub
    
    If DateDiff("d", mobj���ﰲ��.��ʼʱ��, Day) >= 0 And DateDiff("d", Day, mobj���ﰲ��.��ֹʱ��) >= 0 Then
        Metrics.ForeColor = vbBlack
    Else
        '��ǰ���ﰲ��ʱ�䷶Χ����û�ɫ��ʾ������ѡ��
        Metrics.ForeColor = &HC0C0C0
        blnNotSelect = True
    End If
    
    '�ѱ���ļӴ���ʾ
    If CurDayIsSavedPlan(Day) Then
        strStopTxt = CurDayIsNotVisit(Day)
        If strStopTxt <> "" Then
            If IsDate(mobj���ﰲ��(1).��������) Then blnSelected = DateDiff("d", Day, mobj���ﰲ��(1).��������) = 0
            strKey = "B" & IIf(blnSelected, "S", "") & Format(Day, "yyyymmdd")
            If CollExitsValue(mcolPicture, strKey) = False Then Call AddPictureToColl(Day, strStopTxt, blnSelected, True, blnNotSelect)
            If CollExitsValue(mcolPicture, strKey) Then Set Metrics.Picture = mcolPicture(strKey)
        Else
            Metrics.Font.Bold = True
        End If
    Else
        strStopTxt = CurDayIsNotVisit(Day)
        If strStopTxt <> "" Then
            If IsDate(mobj���ﰲ��(1).��������) Then blnSelected = DateDiff("d", Day, mobj���ﰲ��(1).��������) = 0
            strKey = "K" & IIf(blnSelected, "S", "") & Format(Day, "yyyymmdd")
            If CollExitsValue(mcolPicture, strKey) = False Then Call AddPictureToColl(Day, strStopTxt, blnSelected, False, blnNotSelect)
            If CollExitsValue(mcolPicture, strKey) Then Set Metrics.Picture = mcolPicture(strKey)
        Else
            Metrics.Font.Bold = False
        End If
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function CurDayIsSavedPlan(ByVal Day As Date) As Boolean
    Dim i As Integer
    
    '��ǰ�����Ƿ񱣴��˰���
    With mobj���ﰲ��
        If Not .�ѱ�����ﰲ�� Is Nothing Then
            For i = 1 To .�ѱ�����ﰲ��.Count
                If .�ѱ�����ﰲ��(i).�Ƿ�ɾ�� = False Then
                    If IsDate(.�ѱ�����ﰲ��(i).��������) Then
                        If DateDiff("d", Day, .�ѱ�����ﰲ��(i).��������) = 0 Then
                            CurDayIsSavedPlan = True
                            Exit Function
                        End If
                    End If
                End If
            Next
        End If
        
        If Not .δ������ﰲ�� Is Nothing Then
            For i = 1 To .δ������ﰲ��.Count
                If IsDate(.δ������ﰲ��(i).��������) Then
                    If DateDiff("d", Day, .δ������ﰲ��(i).��������) = 0 Then
                        CurDayIsSavedPlan = True
                        Exit Function
                    End If
                End If
            Next
        End If
    End With
End Function

Private Function CurDayIsNotVisit(ByVal Day As Date) As String
    '��ǰ�����Ƿ�ڼ��ջ���ֹͣ����
    '���أ�ֹͣԭ���ڼ�������
    Dim i As Integer
    
    '��ǰ�����Ƿ񱣴��˰���
    With mobj���ﰲ��
        If Not .ͣ���¼ Is Nothing Then
            For i = 1 To .ͣ���¼.Count
                If DateDiff("d", Day, .ͣ���¼(i).��ʼʱ��) <= 0 And DateDiff("d", Day, .ͣ���¼(i).��ֹʱ��) >= 0 Then
                    CurDayIsNotVisit = .ͣ���¼(i).ͣ��ԭ��
                    Exit Function
                End If
            Next
        End If
    End With
End Function

Private Sub AddPictureToColl(ByVal Day As Date, ByVal strSubTxt As String, _
    ByVal blnSelected As Boolean, ByVal blnHavePlan As Boolean, ByVal blnNotSelect As Boolean)
    '���ͼƬ������
    Dim strKey As String, strTxt As String
    Dim objFont As StdFont, objSubFont As StdFont
    Dim lngBackColor As Long, lngForeColor As Long
    Dim objPic As IPictureDisp

    'mcolPicture:ͼƬ���ϣ���ǽڼ��պ�ͣ�ﰲ�ŵġ�
    '���Ϲؼ���(Valueȡ���ڣ���ʽ"yyyymmdd")��
    '   K+Value - δѡ�����δ���ð��ŵ�, KS+Value - ��ǰѡ�����δ���ð��ŵ�
    '   B+Value - δѡ����������˰��ŵ�, BS+Value - ��ǰѡ����������˰��ŵ�
    strKey = IIf(blnHavePlan, "B", "K") & IIf(blnSelected, "S", "") & Format(Day, "yyyymmdd")

    If mcolPicture Is Nothing Then Set mcolPicture = New Collection
    If CollExitsValue(mcolPicture, strKey) = False Then
        Set objFont = New StdFont
        objFont.Name = "����"
        objFont.Size = 9
        objFont.Bold = blnHavePlan
        
        Set objSubFont = New StdFont
        objSubFont.Size = 9
        
        lngBackColor = IIf(blnSelected, dtpDays.PaintManager.SelectedDayBackColor, dtpDays.PaintManager.DayBackColor)
        lngForeColor = IIf(blnNotSelect, &HC0C0C0, vbBlack)
        
        strTxt = Val(Format(Day, "dd"))
        strTxt = IIf(Len(strTxt) = 1, " ", "") & strTxt
        Set objPic = GetTempImage(strTxt, 400, 400, lngBackColor, lngForeColor, objFont, pictxtAlignCenterCenter, _
            Left(strSubTxt, 1), &H5B60F2, objSubFont, pictxtAlignLeftTop)
        If Not objPic Is Nothing Then mcolPicture.Add objPic, strKey
    End If
End Sub

Private Sub dtpDays_SelectionChanged()
    Dim dtCurDate As Date
    Dim blnCancel As Boolean
    Dim strOldDate As String, strNewDate As String
    
    On Error GoTo Errhand
    If mblnNotClick Then Exit Sub
    
    '���ֻ��ѡһ��
    If dtpDays.Selection.BlocksCount > 0 Then
        dtCurDate = dtpDays.Selection.Blocks(0).DateBegin
    Else
        dtpDays.Select mdtMinDay
    End If
    
    strOldDate = mobj���ﰲ��(1).��������: strNewDate = Format(dtCurDate, "yyyy-mm-dd")
    RaiseEvent SelectedChangeBefore(strOldDate, strNewDate, blnCancel)
    If blnCancel Then
        dtpDays.ClearSelection
        dtpDays.Select CDate(strOldDate)
        Exit Sub
    End If
    
    Call ChangeCurPlan(mobj���ﰲ��, strNewDate)
    RaiseEvent SelectedChanged(strOldDate, strNewDate)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub dtpLoopDay_Change()
    On Error GoTo Errhand
    If mblnNotClick Then Exit Sub
    mobj���ﰲ��.��ʼʱ�� = dtpLoopDay.Value
    '�ֶ�����Ϊ���޸ģ���Ϊ��Ȼ����ı��ˣ����Ǽ�¼�����ܻ���һ��
    If mobj���ﰲ��.Count > 0 Then mobj���ﰲ��(1).�Ƿ��޸� = True
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub optLoop_Click(index As Integer)
    Dim blnCancel As Boolean
    Dim strOldDate As String, strNewDate As String
    
    On Error GoTo Errhand
    If mblnNotClick Then Exit Sub
    strOldDate = Val(txtSkip.Text) & "��": strNewDate = Val(txtSkip.Text) & "��"
    RaiseEvent SelectedChangeBefore(strOldDate, strNewDate, blnCancel)
    If blnCancel Then
        mblnNotClick = True
        optLoop(IIf(index = 0, 1, 0)).Value = True
        mblnNotClick = False
        Exit Sub
    End If
    
    dtpLoopDay.Visible = optLoop(0).Value And m_ShowStyle = Show_Plan_Rule
    cboDays.Visible = Not optLoop(0).Value And m_ShowStyle = Show_Plan_Rule
    If index = 0 Then
        mobj���ﰲ��.�Ű���� = 5
        mobj���ﰲ��.��ʼʱ�� = dtpLoopDay.Value
    Else '������ѯ
        mobj���ﰲ��.�Ű���� = 4
        mobj���ﰲ��.��ʼʱ�� = CDate(mDateBefore & cboDays.ItemData(cboDays.ListIndex))
    End If
    Call ChangeCurPlan(mobj���ﰲ��, strNewDate, True)
    '�ֶ�����Ϊ���޸ģ���Ϊ��Ȼ����ı��ˣ����Ǽ�¼�����ܻ���һ��
    If mobj���ﰲ��.Count > 0 Then mobj���ﰲ��(1).�Ƿ��޸� = True
    RaiseEvent SelectedChanged(strOldDate, strNewDate)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

 Private Sub optRule_Click(index As Integer)
    Dim i As Integer, str�������� As String
    Dim bytԭ�Ű���� As Byte
    Dim blnCancel As Boolean
    
    On Error GoTo Errhand
    Call AdjustWeekSeat
    btnWeek.Tag = ""
    If mblnNotClick Then Exit Sub
    
    '�Ű����:1-�����Ű�;2-�����Ű�;3-˫���Ű�;4-������ѭ;5-��ѭ������;6-�ض�����
    bytԭ�Ű���� = mobj���ﰲ��.�Ű����
    mblnNotClick = True
    Select Case index
    Case 0
        btnWeek.ClearAll
        btnWeek.ItemValue(0) = True
        mobj���ﰲ��.�Ű���� = 1
        str�������� = "��һ"
        
        Call SetBoldItem(True)
    Case 1
        mobj���ﰲ��.�Ű���� = 2
        str�������� = "����"
        '����������
        If mobj���ﰲ��.���������� Then chkValied(0).Value = vbChecked
        '���ղ�����
        If mobj���ﰲ��.���ղ����� Then chkValied(1).Value = vbChecked
    Case 2
        mobj���ﰲ��.�Ű���� = 3
        str�������� = "˫��"
        '����������
        If mobj���ﰲ��.���������� Then chkValied(0).Value = vbChecked
        '���ղ�����
        If mobj���ﰲ��.���ղ����� Then chkValied(1).Value = vbChecked
    Case 3
        mobj���ﰲ��.�Ű���� = 4
        txtSkip.Text = "1"
        optLoop(1).Value = True
        mobj���ﰲ��.��ʼʱ�� = CDate(mDateBefore & cboDays.ItemData(cboDays.ListIndex))
        str�������� = Val(txtSkip.Text) & "��"
        '����������
        If mobj���ﰲ��.���������� Then chkValied(0).Value = vbChecked
        '���ղ�����
        If mobj���ﰲ��.���ղ����� Then chkValied(1).Value = vbChecked
    Case 4
        mobj���ﰲ��.�Ű���� = 6
        btnDay.ClearAll
        btnDay.ItemValue(0) = True
        str�������� = 1 & "��"
        
        Call SetBoldItem(False)
    End Select
    Call ChangeCurPlan(mobj���ﰲ��, str��������, bytԭ�Ű���� <> mobj���ﰲ��.�Ű����)
    RaiseEvent SelectedChanged(str��������, str��������)
    mblnNotClick = False
    Exit Sub
Errhand:
    mblnNotClick = False
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub picWeekList_Resize()
    Err = 0: On Error Resume Next
    With picWeekList
        chkValied(0).Top = optRule(4).Top + optRule(4).Height + 100
        chkValied(1).Top = chkValied(0).Top
        picLoop.Left = chkValied(0).Left
        picLoop.Width = .ScaleWidth
        picLoop.Top = chkValied(0).Top + chkValied(0).Height + 50
        lnTop.X1 = -100
        lnTop.X2 = .ScaleWidth + 200
        lnDown.X1 = lnTop.X1
        lnDown.X2 = lnTop.X2
        btnDay.Left = .ScaleLeft  ' chkValied(0).Left
        btnDay.Width = .ScaleWidth - btnDay.Left
        btnDay.Top = lnDown.Y1 + 50
        btnDay.Height = .ScaleHeight - btnDay.Top
        btnWeek.Left = .ScaleLeft
        btnWeek.Width = .ScaleWidth - btnWeek.Left * 2
        If m_ShowStyle = Show_Plan_Rule Then
            btnWeek.Top = chkValied(1).Top
            btnWeek.Height = .ScaleHeight - btnWeek.Top
        Else
            btnWeek.Top = .ScaleTop
            btnWeek.Height = .ScaleHeight
        End If
    End With
End Sub

Private Sub txtSkip_Change()
    Dim blnCancel As Boolean
    Dim strOldDate As String, strNewDate As String
    
    On Error GoTo Errhand
    If mblnNotClick Then Exit Sub
    strOldDate = mobj���ﰲ��(1).��������: strNewDate = Val(txtSkip.Text) & "��"
    RaiseEvent SelectedChangeBefore(strOldDate, strNewDate, blnCancel)
    If blnCancel Then
        mblnNotClick = True
        txtSkip.Text = Val(strOldDate)
        mblnNotClick = False
        Exit Sub
    End If
    
    If optLoop(0).Value Then
        mobj���ﰲ��.�Ű���� = 5
        mobj���ﰲ��.��ʼʱ�� = dtpLoopDay.Value
    Else '������ѯ
        mobj���ﰲ��.�Ű���� = 4
        mobj���ﰲ��.��ʼʱ�� = CDate(mDateBefore & cboDays.ItemData(cboDays.ListIndex))
    End If
    Call ChangeCurPlan(mobj���ﰲ��, strNewDate, True)
    '�ֶ�����Ϊ���޸ģ���Ϊ��Ȼ����ı��ˣ����Ǽ�¼�����ܻ���һ��
    If mobj���ﰲ��.Count > 0 Then mobj���ﰲ��(1).�Ƿ��޸� = True
    RaiseEvent SelectedChanged(strOldDate, strNewDate)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub txtSkip_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub UserControl_Initialize()
    Call InitFace
    Set mcolPicture = New Collection
    Set mobj���ﰲ�� = New ���ﰲ��
End Sub

Private Sub UserControl_Resize()
    Err = 0: On Error Resume Next
    With UserControl
        dtpDays.Left = .ScaleLeft - 300
        dtpDays.Width = .ScaleWidth + 600
        If Not mobj���ﰲ�� Is Nothing Then
            If mobj���ﰲ��.ģ������ = 2 Then
                dtpDays.Top = -550
            Else
                dtpDays.Top = .ScaleTop
            End If
        Else
            dtpDays.Top = .ScaleTop
        End If
        dtpDays.Height = .ScaleHeight - dtpDays.Top
        
        picWeekList.Left = .ScaleLeft
        picWeekList.Width = .ScaleWidth
        picWeekList.Top = .ScaleTop
        picWeekList.Height = .ScaleHeight
    End With
End Sub

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=8,0,0,0
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    SetBackColor Controls, m_BackColor
    dtpDays.PaintManager.ListControlBackColor = m_BackColor
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=8,0,0,0
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
    
    SetEnabled UserControl.Controls, New_Enabled
    btnWeek.Enabled = True
    btnDay.Enabled = True
    dtpDays.Enabled = True
    dtpDays.PaintManager.SelectedDayBackColor = IIf(dtpDays.Enabled, &HFFFFC0, vbRed)
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=6,0,0,0
Public Property Get Font() As Font
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=7,0,0,0
Public Property Get BackStyle() As Integer
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As Integer
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=7,0,0,0
Public Property Get ShowStyle() As gShowStyle
    ShowStyle = m_ShowStyle
End Property

Public Property Let ShowStyle(ByVal New_ShowStyle As gShowStyle)
    m_ShowStyle = New_ShowStyle
    PropertyChanged "ShowStyle"
    Call AdjustWeekSeat
End Property

Public Property Let KeyShift(ByVal Shift As Byte)
    mbytKeyShift = Shift
End Property

'Ϊ�û��ؼ���ʼ������
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_Enabled = m_def_Enabled
    Set m_Font = Ambient.Font
    m_BackStyle = m_def_BackStyle
    m_BorderStyle = m_def_BorderStyle
    m_ShowStyle = m_def_ShowStyle
    Call AdjustWeekSeat
End Sub

'�Ӵ������м�������ֵ
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_ShowStyle = PropBag.ReadProperty("ShowStyle", m_def_ShowStyle)
    
    Call AdjustWeekSeat
    SetBackColor Controls, m_BackColor
End Sub

Private Sub UserControl_Terminate()
    Set mcolPicture = Nothing
    Set mobj���ﰲ�� = Nothing
End Sub

'������ֵд���洢��
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("ShowStyle", m_ShowStyle, m_def_ShowStyle)
End Sub

Private Sub btnDay_Click(idxDay As Integer, Value As Boolean, Text As String)
    Dim strTemp As String, i As Integer
    Dim blnOldSaved As Boolean, blnNewSaved As Boolean
    Dim blnCancel As Boolean
    Dim strOldDate As String, strNewDate As String
    Dim obj�����¼�� As �����¼��
    
    On Error GoTo Errhand
    If mblnNotClick Then Exit Sub
    strOldDate = mobj���ﰲ��(1).��������: strNewDate = btnDay.ItemCaption(idxDay) & "��"
    
    mblnNotClick = True
    If Value = True Then 'ѡ��
        Set obj�����¼�� = mobj���ﰲ��(1).Clone
        
        '��һ��ѡ���Ƿ����ѱ����
        If mobj���ﰲ��.�ѱ�����ﰲ��.Exits("K" & strOldDate) Then
            If mobj���ﰲ��.�ѱ�����ﰲ��("K" & strOldDate).�Ƿ�ɾ�� = False Then blnOldSaved = True
        End If
        If mobj���ﰲ��.δ������ﰲ��.Exits("K" & strOldDate) Then blnOldSaved = True
        
        '����ѡ���Ƿ����ѱ����
        If mobj���ﰲ��.�ѱ�����ﰲ��.Exits("K" & strNewDate) Then
            If mobj���ﰲ��.�ѱ�����ﰲ��("K" & strNewDate).�Ƿ�ɾ�� = False Then blnNewSaved = True
        End If
        If mobj���ﰲ��.δ������ﰲ��.Exits("K" & strNewDate) Then blnNewSaved = True
        
        RaiseEvent SelectedChangeBefore(strOldDate, strNewDate, blnCancel)
        If blnCancel Then
            btnDay.ItemValue(idxDay) = False
            mblnNotClick = False
            Exit Sub
        End If
        
        If blnOldSaved = False And blnNewSaved = False And mbytKeyShift = vbCtrlMask Then
            obj�����¼��.�������� = strNewDate
            mobj���ﰲ��.AddItem obj�����¼��, "K" & strNewDate
        Else
            For i = 0 To btnDay.Count - 1
                btnDay.ItemValue(i) = False
            Next
            btnDay.ItemValue(idxDay) = True
            Call ChangeCurPlan(mobj���ﰲ��, strNewDate)
        End If
        RaiseEvent SelectedChanged(strOldDate, strNewDate)
        Call SetBoldItem(False)
    Else 'ȡ��ѡ��
        If mobj���ﰲ��.�ѱ�����ﰲ��.Exits("K" & strNewDate) Or mobj���ﰲ��.δ������ﰲ��.Exits("K" & strNewDate) _
            Or mobj���ﰲ��.Count = 1 Then
            'ѡ����Ϊ�ѱ����¼����δ�����¼����Ԫ�أ���ǰѡ��ֻ��һ��ʱ��������ȡ��ѡ��
            btnDay.ItemValue(idxDay) = True
        Else
            '�Ƴ�Ԫ��
            If mobj���ﰲ��.Exits("K" & strNewDate) Then mobj���ﰲ��.Remove "K" & strNewDate
            RaiseEvent SelectedChanged(strOldDate, strNewDate)
            Call SetBoldItem(False)
        End If
    End If
    mblnNotClick = False
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
