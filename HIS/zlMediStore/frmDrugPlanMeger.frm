VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmDrugPlanMeger 
   Caption         =   "�ƻ����ϲ�"
   ClientHeight    =   5910
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10350
   Icon            =   "frmDrugPlanMeger.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   10350
   StartUpPosition =   1  '����������
   Begin VB.PictureBox pic�������� 
      Height          =   3120
      Left            =   60
      ScaleHeight     =   3060
      ScaleWidth      =   3435
      TabIndex        =   12
      Top             =   1125
      Width           =   3500
      Begin VB.CommandButton cmdFind 
         Caption         =   "����(&O)"
         Height          =   350
         Left            =   2040
         TabIndex        =   6
         Top             =   2370
         Width           =   1100
      End
      Begin VB.TextBox txt����No 
         Height          =   300
         Left            =   960
         TabIndex        =   2
         Top             =   540
         Width           =   2200
      End
      Begin VB.TextBox txt��ʼNo 
         Height          =   300
         Left            =   960
         TabIndex        =   1
         Top             =   135
         Width           =   2200
      End
      Begin VB.ComboBox cboʱ�䷶Χ 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   960
         Width           =   2200
      End
      Begin MSComCtl2.DTPicker DTP����ʱ�� 
         Height          =   300
         Left            =   960
         TabIndex        =   5
         Top             =   1800
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   120782851
         CurrentDate     =   40750
      End
      Begin MSComCtl2.DTPicker DTP��ʼʱ�� 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "gg yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         Height          =   300
         Left            =   960
         TabIndex        =   4
         Top             =   1380
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   120782851
         CurrentDate     =   40750
      End
      Begin VB.Label lbl����No 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "����NO"
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   540
      End
      Begin VB.Label lbl��ʼNo 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼNO"
         Height          =   180
         Left            =   120
         TabIndex        =   16
         Top             =   180
         Width           =   540
      End
      Begin VB.Label lblʱ�䷶Χ 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��"
         Height          =   180
         Left            =   120
         TabIndex        =   15
         Top             =   1020
         Width           =   720
      End
      Begin VB.Label lbl��ʼʱ�� 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼʱ��"
         Height          =   180
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label lbl����ʱ�� 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��"
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   1860
         Width           =   720
      End
   End
   Begin VB.Frame fra�������� 
      Height          =   495
      Left            =   60
      TabIndex        =   10
      Top             =   495
      Width           =   3500
      Begin VB.Label lbl�������� 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Left            =   120
         TabIndex        =   11
         Top             =   195
         Width           =   720
      End
   End
   Begin VB.Frame fra�ƻ�����Ϣ 
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      Top             =   495
      Width           =   6270
      Begin VB.CheckBox chkAllSelect 
         Caption         =   "ȫѡ"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1425
         TabIndex        =   18
         Top             =   165
         Width           =   975
      End
      Begin VB.Label lbl�ƻ�����Ϣ 
         AutoSize        =   -1  'True
         Caption         =   "�ƻ�����Ϣ"
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   195
         Width           =   900
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   1515
      Left            =   3840
      TabIndex        =   7
      Top             =   1230
      Width           =   6255
      _cx             =   11033
      _cy             =   2672
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
      BackColorSel    =   16777152
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDrugPlanMeger.frx":030A
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
      ExplorerBar     =   1
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
      VirtualData     =   0   'False
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
   Begin VB.Frame fraEW 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   4185
      Left            =   3660
      MousePointer    =   9  'Size W E
      TabIndex        =   9
      Top             =   465
      Width           =   45
   End
   Begin XtremeCommandBars.CommandBars comBars 
      Left            =   315
      Top             =   90
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager imgIcon 
      Bindings        =   "frmDrugPlanMeger.frx":037F
      Left            =   1515
      Top             =   90
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmDrugPlanMeger.frx":0393
   End
End
Attribute VB_Name = "frmDrugPlanMeger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnChkFocus As Boolean
Private mblnChang    As Boolean

Private Const MCONMEGER = 2 '�ϲ�
Private Const MCONEXIT = 4  '�˳�

Private Const MCONѡ�� = 0
Private Const MCONNO = 1
Private Const MCONID = 2
Private Const MCON�ƻ����� = 3
Private Const MCON�ڼ� = 4
Private Const MCON�����ⷿ = 5
Private Const MCON���Ʒ��� = 6
Private Const MCON������ = 7
Private Const MCON�������� = 8
Private Const MCON����� = 9
Private Const MCON������� = 10
Private Const MCON������ = 11
Private Const MCON�������� = 12
Private Const MCON����˵�� = 13
Private Const MCONROWS = 14

Private Sub GetList()
'��ȡ�ƻ�������
    Dim cmdControl As CommandBarControl
    Dim rsList As New Recordset
    Dim strFind As String
    Dim lngRow As Long
     
    On Error GoTo errHandle
    
    With vsfList
        If Me.txt��ʼNo <> "" And Me.txt����No <> "" Then strFind = " And A.No >= [3] And A.No <=[4] "
        If Me.txt��ʼNo <> "" And Me.txt����No = "" Then strFind = " And A.No >= [3] "
        If Me.txt��ʼNo = "" And Me.txt����No <> "" Then strFind = " And A.No <= [4] "
        
        gstrSQL = " SELECT '' as ѡ��, a.NO, a.ID, DECODE(a.�ƻ�����,0,'��ʱ',1,'�¶ȼƻ�',2,'���ȼƻ�',3,'��ȼƻ�','�ܼƻ�') AS �ƻ����� ," & _
            "a.�ڼ�, b.���� as �����ⷿ, DECODE(A.���Ʒ���, 0, '�����������', 1, '����ͬ�����β��շ�', 2, '�ٽ��ڼ�ƽ�����շ�', 3, 'ҩƷ����������շ�', 4, 'ҩƷ�����������շ�', '�Զ���������շ�') AS ���Ʒ��� ," & _
            "a.������,TO_CHAR(a.��������,'YYYY-MM-DD HH24:MI:SS') AS ��������, a.�����, " & _
            "TO_CHAR(a.�������,'YYYY-MM-DD HH24:MI:SS') AS �������,a.������,TO_CHAR(a.��������,'YYYY-MM-DD HH24:MI:SS') AS ��������, a.����˵�� " & _
            " FROM ҩƷ�ɹ��ƻ� A, ���ű� B " & _
            " WHERE a.�ⷿID= b.ID(+) And NVL(a.ҩ��id,0)=0 And NVL(a.�ⷿID,0)<>0 And NVL(a.�ϲ��ƻ�id,0)=0 And a.�������� between [1] And to_date([2],'YYYY-MM-DD HH24:MI:SS') " & strFind & _
            " ORDER BY A.NO DESC "
            
        Set rsList = zlDataBase.OpenSQLRecord(gstrSQL, "�ϲ��ƻ���", DTP��ʼʱ��.Value, Format(DTP����ʱ��.Value, "yyyy-mm-dd") & " 23:59:59", txt��ʼNo.Text, txt����No.Text)
        .rows = 1
        
        If rsList.EOF Then
            .rows = 2
            .Editable = flexEDNone
            chkAllSelect.Enabled = False
        Else
            .rows = rsList.RecordCount + 1
            .Editable = flexEDKbdMouse
            chkAllSelect.Enabled = True
            For lngRow = 1 To .rows - 1
                .TextMatrix(lngRow, .ColIndex("NO")) = rsList!NO
                .TextMatrix(lngRow, .ColIndex("ID")) = rsList!id
                .TextMatrix(lngRow, .ColIndex("�ƻ�����")) = rsList!�ƻ�����
                .TextMatrix(lngRow, .ColIndex("�ڼ�")) = rsList!�ڼ�
                .TextMatrix(lngRow, .ColIndex("�����ⷿ")) = rsList!�����ⷿ
                .TextMatrix(lngRow, .ColIndex("���Ʒ���")) = rsList!���Ʒ���
                .TextMatrix(lngRow, .ColIndex("������")) = rsList!������
                .TextMatrix(lngRow, .ColIndex("��������")) = rsList!��������
                .TextMatrix(lngRow, .ColIndex("�����")) = rsList!�����
                .TextMatrix(lngRow, .ColIndex("�������")) = rsList!�������
                .TextMatrix(lngRow, .ColIndex("������")) = rsList!������
                .TextMatrix(lngRow, .ColIndex("��������")) = rsList!��������
                .TextMatrix(lngRow, .ColIndex("����˵��")) = IIf(IsNull(rsList!����˵��), "", rsList!����˵��)
                rsList.MoveNext
            Next
        End If
        
        .Row = 1
        .SetFocus
        chkAllSelect.Value = 0
        rsList.Close
    End With
    
    Set cmdControl = comBars.FindControl(, MCONMEGER)
    cmdControl.Enabled = False
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub MegerPlan()
'�ϲ��ƻ���
    Dim intRow As Integer
    Dim strID As String
    Dim strNewNO As String
    Dim lngNewId As Long
    Dim arrSql As Variant
    Dim i As Integer
    
    On Error GoTo errHandle
    
    arrSql = Array()
    strNewNO = Sys.GetNextNo(32)
    lngNewId = Sys.NextId("ҩƷ�ɹ��ƻ�")
    strID = lngNewId & "|"
    With vsfList
        For intRow = 1 To .rows - 1
            If Val(.TextMatrix(intRow, .ColIndex("ѡ��"))) = -1 Then
                strID = strID & .TextMatrix(intRow, .ColIndex("id")) & ","
            End If
        Next
    End With
    
    gstrSQL = "zl_ҩƷ�ƻ���������_INSERT("
        '�ƻ�ID
        gstrSQL = gstrSQL & lngNewId
        'NO
        gstrSQL = gstrSQL & ",'" & strNewNO & "'"
        '�ƻ�����
        gstrSQL = gstrSQL & ",1"
        '�ڼ�
        gstrSQL = gstrSQL & ",'" & Format(DateAdd("m", 1, Sys.Currentdate), "yyyyMM") & "'"
        '�ⷿID
        gstrSQL = gstrSQL & ",Null"
        'ҩ��ID
        gstrSQL = gstrSQL & ",Null"
        '���Ʒ���
        gstrSQL = gstrSQL & ",1"
        '������
        gstrSQL = gstrSQL & ",'" & UserInfo.�û����� & "'"
        '��������
        gstrSQL = gstrSQL & ",to_date('" & Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS')"
        '����˵��
        gstrSQL = gstrSQL & ",''"
        '��Դ�ⷿ
        gstrSQL = gstrSQL & ",'0'"
        gstrSQL = gstrSQL & ")"
        
    ReDim Preserve arrSql(UBound(arrSql) + 1)
    arrSql(UBound(arrSql)) = gstrSQL
    
    gstrSQL = "Zl_ҩƷ�ƻ�����_Union('" & Mid(strID, 1, Len(strID) - 1) & "')"
    ReDim Preserve arrSql(UBound(arrSql) + 1)
    arrSql(UBound(arrSql)) = gstrSQL
    
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSql)
        Call zlDataBase.ExecuteProcedure(CStr(arrSql(i)), "MegerPlan")
    Next
    gcnOracle.CommitTrans
    
    '���»�ȡ����
    Call GetList
    MsgBox "�ϲ��ɹ���", vbInformation, gstrSysName
    
    Exit Sub
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetListCol()
'�����������
    Dim intCol As Integer
    
    With vsfList
        .rows = 2
        .Cols = MCONROWS
        .ColDataType(0) = flexDTBoolean
        .Editable = flexEDNone
        .TextMatrix(0, MCONѡ��) = "ѡ��"
        .TextMatrix(0, MCONNO) = "NO"
        .TextMatrix(0, MCONID) = "ID"
        .TextMatrix(0, MCON�ƻ�����) = "�ƻ�����"
        .TextMatrix(0, MCON�ڼ�) = "�ڼ�"
        .TextMatrix(0, MCON�����ⷿ) = "�����ⷿ"
        .TextMatrix(0, MCON���Ʒ���) = "���Ʒ���"
        .TextMatrix(0, MCON������) = "������"
        .TextMatrix(0, MCON��������) = "��������"
        .TextMatrix(0, MCON�����) = "�����"
        .TextMatrix(0, MCON�������) = "�������"
        .TextMatrix(0, MCON������) = "������"
        .TextMatrix(0, MCON��������) = "��������"
        .TextMatrix(0, MCON����˵��) = "����˵��"
        
        For intCol = 0 To .Cols - 1
            .ColKey(intCol) = .TextMatrix(0, intCol)
        Next
        .ColWidth(.ColIndex("ѡ��")) = 500
        .ColWidth(.ColIndex("�����ⷿ")) = 1200
        .ColWidth(.ColIndex("���Ʒ���")) = 1800
        .ColWidth(.ColIndex("��������")) = 1000
        .ColWidth(.ColIndex("�������")) = 1000
        .ColWidth(.ColIndex("��������")) = 1000
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .ColAlignment(.ColIndex("NO")) = flexAlignLeftCenter
        .ColAlignment(.ColIndex("�ڼ�")) = flexAlignLeftCenter
        .ColWidth(.ColIndex("ID")) = 0
        
    End With
End Sub

Public Sub showMe(ByVal frmPar As Form)
    Me.Show vbModal, frmPar
End Sub

Private Sub chkAllSelect_Click()
    Dim i As Integer
    
    With vsfList
        If mblnChkFocus Then
            For i = 1 To .rows - 1
                .TextMatrix(i, .ColIndex("ѡ��")) = IIf(chkAllSelect.Value = 1, -1, 0)
            Next
        End If
    End With
End Sub

Private Sub chkAllSelect_GotFocus()
    mblnChkFocus = True
End Sub

Private Sub chkAllSelect_LostFocus()
    mblnChkFocus = False
End Sub

Private Sub cmdFind_Click()
    Call GetList
End Sub

Private Sub combars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.id
        Case MCONMEGER
            Call MegerPlan
        Case MCONEXIT
            Call CheckUnLoad
    End Select
End Sub

Private Sub txt��ʼNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Len(txt��ʼNo) < 8 And Len(txt��ʼNo) > 0 Then
            txt��ʼNo.Text = zlCommFun.GetFullNO(txt��ʼNo.Text, 32)
        End If
        Me.txt����No.SetFocus
    End If
End Sub

Private Sub txt��ʼNo_LostFocus()
    If Len(txt��ʼNo) < 8 And Len(txt��ʼNo) > 0 Then
        txt��ʼNo.Text = zlCommFun.GetFullNO(txt��ʼNo.Text, 32)
    End If
End Sub

Private Sub txt����NO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Len(txt����No) < 8 And Len(txt����No) > 0 Then
            txt����No.Text = zlCommFun.GetFullNO(txt����No.Text, 32)
        End If
        Me.cboʱ�䷶Χ.SetFocus
    End If
End Sub

Private Sub txt����No_LostFocus()
    If Len(txt����No) < 8 And Len(txt����No) > 0 Then
        txt����No.Text = zlCommFun.GetFullNO(txt����No.Text, 32)
    End If
End Sub

Private Sub cboʱ�䷶Χ_Click()
    If Me.cboʱ�䷶Χ.ListIndex = 0 Then
        Me.DTP��ʼʱ��.Value = Date
        Me.DTP����ʱ��.Value = Date
        Me.DTP��ʼʱ��.Enabled = False
        Me.DTP����ʱ��.Enabled = False
    ElseIf Me.cboʱ�䷶Χ.ListIndex = 1 Then
        Me.DTP��ʼʱ��.Value = Date - 1
        Me.DTP����ʱ��.Value = Date
        Me.DTP��ʼʱ��.Enabled = False
        Me.DTP����ʱ��.Enabled = False
    ElseIf Me.cboʱ�䷶Χ.ListIndex = 2 Then
        Me.DTP��ʼʱ��.Value = Date - 2
        Me.DTP����ʱ��.Value = Date
        Me.DTP��ʼʱ��.Enabled = False
        Me.DTP����ʱ��.Enabled = False
    ElseIf Me.cboʱ�䷶Χ.ListIndex = 3 Then
        Me.DTP��ʼʱ��.Value = Date - 6
        Me.DTP����ʱ��.Value = Date
        Me.DTP��ʼʱ��.Enabled = False
        Me.DTP����ʱ��.Enabled = False
    Else
        Me.DTP��ʼʱ��.Value = Date - 30
        Me.DTP����ʱ��.Value = Date
        Me.DTP��ʼʱ��.Enabled = True
        Me.DTP����ʱ��.Enabled = True
    End If
End Sub

Private Sub cboʱ�䷶Χ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call OS.PressKey(vbKeyTab)
    End If
End Sub

Private Sub dtp��ʼʱ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call OS.PressKey(vbKeyTab)
    End If
End Sub

Private Sub dtp����ʱ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call OS.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    Call InitComman
    Call InitTool
    Call InitCbo
    Call SetListCol
End Sub

Private Sub InitCbo()
'����������
    With cboʱ�䷶Χ
        .Clear
        .AddItem "һ����"
        .AddItem "������"
        .AddItem "������"
        .AddItem "һ����"
        .AddItem "�Զ���ʱ��"
        .ListIndex = 0
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    fra��������.Move 10, , 3500
    fraEW.Move fra��������.Left + fra��������.Width, fra��������.Top, 45, Me.ScaleHeight
    fra�ƻ�����Ϣ.Move fraEW.Left + fraEW.Width, fra��������.Top, Me.ScaleWidth - fraEW.Left - fraEW.Width - 20
    pic��������.Move fra��������.Left, fra��������.Top + fra��������.Height + 10, fra��������.Width, Me.ScaleHeight - fra��������.Top - fra��������.Height - 55
    vsfList.Move fra�ƻ�����Ϣ.Left, fra�ƻ�����Ϣ.Top + fra�ƻ�����Ϣ.Height + 10, fra�ƻ�����Ϣ.Width, Me.ScaleHeight - fra�ƻ�����Ϣ.Top - fra�ƻ�����Ϣ.Height - 50
End Sub

Private Sub fraEW_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'------------------------------------------
'�������͵�����������
'------------------------------------------
    On Error Resume Next
    If Me.fra��������.Width + x < 3000 Or Me.vsfList.Width - x < 4000 Then
        Exit Sub
    End If
    
    If Button = 1 Then
        Me.fraEW.Move Me.fraEW.Left + x, Me.fraEW.Top, Me.fraEW.Width, Me.fraEW.Height
        Me.fra��������.Move Me.fra��������.Left, Me.fra��������.Top, Me.fra��������.Width + x, Me.fra��������.Height
        Me.fra�ƻ�����Ϣ.Move Me.fra�ƻ�����Ϣ.Left + x, Me.fra�ƻ�����Ϣ.Top, Me.fra�ƻ�����Ϣ.Width - x, Me.fra�ƻ�����Ϣ.Height
        
        Me.pic��������.Move Me.pic��������.Left, Me.pic��������.Top, Me.pic��������.Width + x, Me.pic��������.Height
        Me.vsfList.Move Me.vsfList.Left + x, Me.vsfList.Top, Me.vsfList.Width - x, Me.vsfList.Height
        Me.cmdFind.Move cmdFind.Left + x
        
        Me.txt����No.Width = Me.txt����No.Width + x
        Me.txt��ʼNo.Width = Me.txt��ʼNo.Width + x
        Me.cboʱ�䷶Χ.Width = Me.cboʱ�䷶Χ.Width + x
        Me.DTP����ʱ��.Width = Me.DTP����ʱ��.Width + x
        Me.DTP��ʼʱ��.Width = Me.DTP��ʼʱ��.Width + x
    End If
End Sub

Private Sub InitComman()
'--------------------------------------
'��ʼ��CommandBars�ؼ�
'--------------------------------------
    With CommandBarsGlobalSettings
        Set .App = App
        .ResourceFile = .OcxPath & "\XTPResourceZhCn.dll" '��������������Դ�ļ�
        .ColorManager.SystemTheme = xtpSystemThemeAuto  '�ؼ��������ɫ����������ϵͳ�Զ�ʶ��
    End With

    With comBars.Options
        .ShowExpandButtonAlways = False '�����ڹ������Ҳ���ʾѡ�ť,��ʹ�������㹻��
        .ToolBarAccelTips = True '��ʾ��ť��ʾ
        .AlwaysShowFullMenus = False '�����õĲ˵���������
        .UseFadedIcons = True 'ͼ����ʾΪ��ɫЧ��
        .IconsWithShadow = True '���ָ�������ͼ����ʾ��ӰЧ��
        .UseDisabledIcons = True '��������ť����ʱͼ����ʾΪ������ʽ
        .LargeIcons = True '��������ʾΪ��ͼ��
        .SetIconSize True, 24, 24 '���ô�ͼ��ĳߴ�
        .SetIconSize False, 16, 16 '����Сͼ��ĳߴ�
    End With

    With comBars
        .VisualTheme = xtpThemeOffice2003 '���ÿؼ���ʾ���
        .EnableCustomization False '�Ƿ������Զ�������
        .Item(1).Delete
        .Icons = imgIcon.Icons
    End With
End Sub

Private Sub InitTool()
'-----------------------------------------------------
'���ù�����
'----------------------------------------------------
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    
    Set objBar = comBars.Add("������1", xtpBarTop)
    objBar.ContextMenuPresent = False '�������ϵ������Ҽ�ʱ���������ò˵�
    objBar.ShowTextBelowIcons = False '�������еİ�ť������ʾ��ͼ���Ҳ�
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, MCONMEGER, "�ϲ�")
        objControl.Style = xtpButtonIconAndCaption
        objControl.Enabled = False
        Set objControl = .Add(xtpControlButton, MCONEXIT, "�˳�")
        objControl.Style = xtpButtonIconAndCaption
        objControl.BeginGroup = True
    End With
End Sub

Private Sub vsfList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfList
        If Col <> 0 Then
            Cancel = True
        End If
        mblnChkFocus = False
    End With
End Sub

Private Sub vsfList_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim cmdControl As CommandBarControl
    Dim intRow As Integer
    Dim intCount As Integer
    
    With vsfList
        If Row > 0 And Val(.TextMatrix(Row, 2)) <> 0 And Col = 0 Then
            For intRow = 1 To .rows - 1
                If Val(.TextMatrix(intRow, 0)) = -1 Then
                    intCount = intCount + 1
                End If
            Next
            
            '�Ƿ�ѡ����������
            Set cmdControl = comBars.FindControl(, MCONMEGER)
            If intCount >= 2 Then
                cmdControl.Enabled = True
            Else
                cmdControl.Enabled = False
            End If
            
            '�Ƿ�ȫѡ
            If mblnChkFocus = False Then
                If intCount = .rows - 1 Then
                    chkAllSelect.Value = 1
                Else
                    chkAllSelect.Value = 0
                End If
            End If
        End If
    End With
End Sub

Private Sub vsfList_ChangeEdit()
    mblnChang = True
End Sub

Private Sub CheckUnLoad()
'�˳�ǰ����Ƿ���ѡ�еĵ���
    Dim intRow As Integer
    Dim blnChanged As Boolean
    
    blnChanged = True
    With vsfList
        For intRow = 1 To .rows - 1
            If Val(.TextMatrix(intRow, 0)) = -1 Then
                If MsgBox("������ѡ�еļƻ�����δ�ϲ����Ƿ�ȷ���˳���", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo Then
                    blnChanged = False
                End If
                Exit For
            End If
        Next
            
        If blnChanged Then Unload Me
    End With
End Sub

