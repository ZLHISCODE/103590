VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmClinicPlanChangeHistory 
   Caption         =   "�ٴ����ﰲ�ű䶯��Ϣ"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11760
   Icon            =   "frmClinicPlanChangeHistory.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   11760
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picButton 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   11760
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   7695
      Width           =   11760
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "�˳�(&E)"
         Height          =   350
         Left            =   10230
         TabIndex        =   9
         Top             =   180
         Width           =   1100
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "����(&H)"
         Height          =   350
         Left            =   450
         TabIndex        =   8
         Top             =   180
         Width           =   1100
      End
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "��ѯ(&F)"
      Height          =   350
      Left            =   5760
      TabIndex        =   4
      Top             =   60
      Width           =   1100
   End
   Begin MSComCtl2.DTPicker dtpStartDate 
      Height          =   300
      Left            =   900
      TabIndex        =   1
      Top             =   90
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   529
      _Version        =   393216
      CalendarTitleBackColor=   -2147483630
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   172556291
      CurrentDate     =   42453
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfChangeInfo 
      Height          =   6735
      Left            =   30
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   450
      Width           =   10245
      _cx             =   18071
      _cy             =   11880
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483638
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmClinicPlanChangeHistory.frx":6852
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
      Begin VB.PictureBox picImgPlan 
         BorderStyle     =   0  'None
         Height          =   135
         Left            =   30
         Picture         =   "frmClinicPlanChangeHistory.frx":68C7
         ScaleHeight     =   135
         ScaleWidth      =   150
         TabIndex        =   6
         Top             =   60
         Width           =   150
      End
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   300
      Left            =   3330
      TabIndex        =   3
      Top             =   90
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   529
      _Version        =   393216
      CalendarTitleBackColor=   -2147483630
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   172556291
      CurrentDate     =   42453
   End
   Begin VB.Label lblTimeRange 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      Height          =   180
      Left            =   3090
      TabIndex        =   2
      Top             =   150
      Width           =   180
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ѯʱ��"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   720
   End
End
Attribute VB_Name = "frmClinicPlanChangeHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModule As Long
Private mblnFirst As Boolean

Public Function ShowMe(frmParent As Object, ByVal lngModule As Long) As Boolean
    '�������
    mlngModule = lngModule
    Err = 0: On Error Resume Next
    
    Me.Show 1, frmParent
End Function

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    If DateDiff("s", dtpStartDate.Value, dtpEndDate.Value) <= 0 Then
        MsgBox "��ѯ��ֹʱ�������ڿ�ʼʱ�䣡", vbInformation, gstrSysName
        If dtpEndDate.Visible And dtpEndDate.Enabled Then dtpEndDate.SetFocus
        Exit Sub
    End If
    Call RefreshData
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub dtpEndDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: Exit Sub
End Sub

Private Sub dtpStartDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: Exit Sub
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    If cmdFind.Visible And cmdFind.Enabled Then cmdFind.SetFocus
End Sub

Private Sub Form_Load()
    mblnFirst = True
    '��ʼ�����ڣ�Ĭ����ʾһ������
    dtpStartDate.Value = Format(Now - 7, "yyyy-mm-dd hh:mm:ss")
    dtpEndDate.Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
    
    '��ʼ�����
    Call InitGrid
    Call zl_vsGrid_Para_Restore(mlngModule, vsfChangeInfo, Me.Name, "�䶯��Ϣ")
    
    Call RefreshData
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With vsfChangeInfo
        .Left = 20
        .Top = 450
        .Width = Me.ScaleWidth - .Left * 2
        .Height = Me.ScaleHeight - picButton.Height - .Top - 20
    End With
End Sub

Private Sub InitGrid()
    '��ʼ�����
    Dim strHead As String, varData As Variant
    Dim strHeadSub As String, varDataSub As Variant
    Dim i As Long, lngCol As Long
    Dim arrDate As Variant
    Dim dtCurDate As Date, intDays As Integer
    
    Err = 0: On Error GoTo errHandler
    With vsfChangeInfo
        .Redraw = False
        .Rows = 2
        
        strHead = ",4,220|����,4,500|����,4,500|����,1,1000|��Ŀ,1,0|ҽ��,1,700|�䶯����,4,1100|" & _
                "�䶯ԭ��,1,1200|�䶯ǰ����,1,2500|�䶯������,1,2500|" & _
                "�Ǽ���,1,700|�Ǽ�ʱ��,4,1900|������,1,700|����ʱ��,4,1900|ȡ����,1,700|ȡ��ʱ��,4,1900"
        varData = Split(strHead, "|")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .ColKey(i) = Split(varData(i), ",")(0)
        Next
        .FixedCols = 1: .FixedRows = 1
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .ExplorerBar = flexExSort
        .HighLight = flexHighlightAlways
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .GridLines = flexGridFlat
        '.WordWrap = True '�����Զ�����
        .RowHeightMin = 350
        
        '����������,�����û�ѡ����ʾ��
        For i = 0 To .Cols - 1
            'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)|������(0-��������,1-��ֹ����,2-��������,�����س���������)
            Select Case i
            Case .ColIndex("����"), .ColIndex("����"), .ColIndex("ҽ��")
                .ColData(i) = "1|0"
            End Select
        Next
        .Redraw = True
    End With
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function RefreshData() As Boolean
    Dim strSQL As String, lngRow As Long
    Dim rsData As ADODB.Recordset
    Dim dtStart As Date, dtEnd As Date
    
    Err = 0: On Error GoTo errHandler
    dtStart = dtpStartDate.Value
    dtEnd = dtpEndDate.Value
    
    vsfChangeInfo.Clear 1
    vsfChangeInfo.Rows = 2
    zlCommFun.ShowFlash "���ڼ������ݣ����Ե�...", Me
    '�޺š���Լ�����ҵ���
    strSQL = "Select Max(��¼id) As ��¼id, Max(�䶯ԭ��) As �䶯ԭ��, Max(�䶯ǰ || Decode(�䶯����, 1, ����)) As �䶯ǰ," & vbNewLine & _
            "        Max(�䶯�� || Decode(�䶯����, 2, ����)) As �䶯��, Max(�Ǽ���) As �Ǽ���, Max(�Ǽ�ʱ��) As �Ǽ�ʱ��, Max(�Ǽ���) As ������, Max(�Ǽ�ʱ��) As ����ʱ��," & vbNewLine & _
            "        Null As ȡ����, Null As ȡ��ʱ��" & vbNewLine & _
            " From (Select m.ID As �䶯id, n.�䶯����, Max(m.��¼id) As ��¼id, Max(Decode(m.�䶯����, 1, '�޺ŵ���', 2, '��Լ����', 3, '���ұ䶯')) As �䶯ԭ��," & vbNewLine & _
            "               Max(Decode(m.�䶯����, 1, '�޺�:' || ԭ����, 2, '��Լ:' || ԭ����, 3," & vbNewLine & _
            "                           Decode(ԭ���﷽ʽ, 0, '������', 1, 'ָ������:', 2, '��̬����:', 3, 'ƽ������:'))) As �䶯ǰ," & vbNewLine & _
            "               Max(Decode(m.�䶯����, 1, '�޺�:' || ������, 2, '��Լ:' || ������, 3," & vbNewLine & _
            "                           Decode(�ַ��﷽ʽ, 0, '������', 1, 'ָ������:', 2, '��̬����:', 3, 'ƽ������:'))) As �䶯��," & vbNewLine & _
            "               f_List2str(Cast(Collect(n.��������) As t_Strlist)) As ����, Max(m.����Ա����) As �Ǽ���, Max(m.�Ǽ�ʱ��) As �Ǽ�ʱ��" & vbNewLine & _
            "        From �ٴ�����䶯��¼ M, �ٴ�����䶯��ϸ N" & vbNewLine & _
            "        Where m.Id = n.�䶯id(+) And m.�䶯���� In (1, 2, 3) And m.�Ǽ�ʱ�� Between [1] And [2]" & vbNewLine & _
            "        Group By m.id, n.�䶯����)" & vbNewLine & _
            " Group By �䶯id"
    'ͣ�����
    strSQL = strSQL & vbNewLine & _
            " Union All" & vbNewLine & _
            " Select m.��¼id, Decode(m.����ҽ������, Null, 'ͣ��', '����') As �䶯ԭ��, '' As �䶯ǰ," & vbNewLine & _
            "       Decode(m.����ҽ������, Null," & vbNewLine & _
            "               'ͣ��ʱ��:' || To_Char(n.��������, 'yyyy-mm-dd') || ' ' || To_Char(m.��ʼʱ��, 'hh24:mi') || '��' ||" & vbNewLine & _
            "                To_Char(m.��ֹʱ��, 'hh24:mi'), n.�ϰ�ʱ�� || ',����ҽ��:' || m.����ҽ������) As �䶯��, m.������ As �Ǽ���, m.����ʱ�� As �Ǽ�ʱ��," & vbNewLine & _
            "       m.������, m.����ʱ��, m.ȡ����, m.ȡ��ʱ��" & vbNewLine & _
            " From �ٴ�����ͣ���¼ M, �ٴ������¼ N" & vbNewLine & _
            " Where m.��¼id Is Not Null And m.��¼id = n.Id And m.����ʱ�� Is Not Null And m.����ʱ�� Between [1] And [2]"
    
    strSQL = "Select c.����, c.����, d.���� As ��Ŀ, e.���� As ����, b.ҽ������ As ҽ��, To_Char(b.��������, 'yyyy-mm-dd') As �䶯����," & _
            "        a.�䶯ԭ��, a.�䶯ǰ, a.�䶯��, a.�Ǽ���, To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��, " & vbNewLine & _
            "        a.�Ǽ��� As ������, To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, " & vbNewLine & _
            "        a.ȡ����, Decode(a.ȡ��ʱ��, Null, '', To_Char(a.ȡ��ʱ��, 'yyyy-mm-dd hh24:mi:ss')) As ȡ��ʱ��" & vbNewLine & _
            " From (" & strSQL & ") A, �ٴ������¼ B," & vbNewLine & _
            "      �ٴ������Դ C, �շ���ĿĿ¼ D, ���ű� E" & vbNewLine & _
            " Where a.��¼id = b.Id And b.��Դid = c.Id And b.��Ŀid = d.Id And b.����id = e.Id(+)"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dtStart, dtEnd)
    If rsData Is Nothing Then Exit Function
    If rsData.RecordCount = 0 Then Exit Function
    
    '��������
    With vsfChangeInfo
        .Redraw = False
        .Rows = rsData.RecordCount + 1
        lngRow = 1
        Do While Not rsData.EOF
            .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsData!����)
            .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsData!����)
            .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsData!����)
            .TextMatrix(lngRow, .ColIndex("��Ŀ")) = Nvl(rsData!��Ŀ)
            .TextMatrix(lngRow, .ColIndex("ҽ��")) = Nvl(rsData!ҽ��)
            .TextMatrix(lngRow, .ColIndex("�䶯����")) = Nvl(rsData!�䶯����)
            .TextMatrix(lngRow, .ColIndex("�䶯ԭ��")) = Nvl(rsData!�䶯ԭ��)
            .TextMatrix(lngRow, .ColIndex("�䶯ǰ����")) = Nvl(rsData!�䶯ǰ)
            .TextMatrix(lngRow, .ColIndex("�䶯������")) = Nvl(rsData!�䶯��)
            .TextMatrix(lngRow, .ColIndex("�Ǽ���")) = Nvl(rsData!�Ǽ���)
            .TextMatrix(lngRow, .ColIndex("�Ǽ�ʱ��")) = Nvl(rsData!�Ǽ�ʱ��)
            .TextMatrix(lngRow, .ColIndex("������")) = Nvl(rsData!������)
            .TextMatrix(lngRow, .ColIndex("����ʱ��")) = Nvl(rsData!����ʱ��)
            .TextMatrix(lngRow, .ColIndex("ȡ����")) = Nvl(rsData!ȡ����)
            .TextMatrix(lngRow, .ColIndex("ȡ��ʱ��")) = Nvl(rsData!ȡ��ʱ��)
            lngRow = lngRow + 1
            rsData.MoveNext
        Loop
        .Redraw = True
    End With
    zlCommFun.StopFlash
    RefreshData = True
    Exit Function
errHandler:
    zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
     Call zl_vsGrid_Para_Save(mlngModule, vsfChangeInfo, Me.Name, "�䶯��Ϣ")
End Sub

Private Sub picButton_Resize()
    On Error Resume Next
    cmdExit.Left = picButton.ScaleWidth - cmdExit.Width - 500
End Sub

Private Sub vsfChangeInfo_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call zl_vsGrid_Para_Save(mlngModule, vsfChangeInfo, Me.Name, "�䶯��Ϣ")
End Sub

Private Sub vsfChangeInfo_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub picImgPlan_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    
    vRect = zlControl.GetControlRect(picImgPlan.Hwnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgPlan.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsfChangeInfo, lngLeft, lngTop, picImgPlan.Height)
    Call zl_vsGrid_Para_Save(mlngModule, vsfChangeInfo, Me.Name, "�䶯��Ϣ")
End Sub
