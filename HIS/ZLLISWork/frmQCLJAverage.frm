VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~3.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "CODEJO~1.OCX"
Begin VB.Form frmQCLJAverage 
   Caption         =   "��ֵLJ�ʿز�ѯ"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11490
   Icon            =   "frmQCLJAverage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   11490
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picRecord 
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      Height          =   6750
      Left            =   75
      ScaleHeight     =   6750
      ScaleWidth      =   2445
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   540
      Width           =   2445
      Begin VB.CommandButton cmdˢ�� 
         Height          =   600
         Left            =   2085
         Picture         =   "frmQCLJAverage.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   90
         Width           =   330
      End
      Begin MSComCtl2.DTPicker dtp���� 
         Height          =   300
         Index           =   0
         Left            =   435
         TabIndex        =   8
         Top             =   75
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   206766083
         CurrentDate     =   39110
      End
      Begin MSComCtl2.DTPicker dtp���� 
         Height          =   300
         Index           =   1
         Left            =   435
         TabIndex        =   9
         Top             =   390
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   206766083
         CurrentDate     =   39110
      End
      Begin VSFlex8Ctl.VSFlexGrid vfgItem 
         Height          =   4830
         Left            =   45
         TabIndex        =   10
         Top             =   735
         Width           =   2445
         _cx             =   4313
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
         BackColorFixed  =   14737632
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16635590
         ForeColorSel    =   -2147483640
         BackColorBkg    =   14737632
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
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
      Begin VB.Label lbl���� 
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Index           =   0
         Left            =   45
         TabIndex        =   12
         Top             =   135
         Width           =   405
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Index           =   1
         Left            =   210
         TabIndex        =   11
         Top             =   420
         Width           =   180
      End
   End
   Begin VB.ComboBox cbo���� 
      Height          =   300
      Left            =   2070
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   90
      Width           =   1845
   End
   Begin VB.ComboBox cbo���� 
      Height          =   300
      Left            =   4905
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   75
      Width           =   2115
   End
   Begin VB.PictureBox picCharts 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4395
      Left            =   4890
      ScaleHeight     =   4395
      ScaleWidth      =   6510
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   495
      Width           =   6510
      Begin XtremeSuiteControls.TabControl tbcCharts 
         Height          =   3975
         Left            =   150
         TabIndex        =   2
         Top             =   165
         Width           =   6105
         _Version        =   589884
         _ExtentX        =   10769
         _ExtentY        =   7011
         _StockProps     =   64
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7605
      Width           =   11490
      _ExtentX        =   20267
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmQCLJAverage.frx":6BDC
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15187
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
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
   Begin C1Chart2D8.Chart2D chtCopy 
      Height          =   435
      Left            =   1260
      TabIndex        =   5
      Top             =   60
      Visible         =   0   'False
      Width           =   765
      _Version        =   524288
      _Revision       =   7
      _ExtentX        =   1349
      _ExtentY        =   767
      _StockProps     =   0
      ControlProperties=   "frmQCLJAverage.frx":746E
   End
   Begin VB.PictureBox picData 
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   3840
      ScaleHeight     =   1815
      ScaleWidth      =   5280
      TabIndex        =   13
      Top             =   5040
      Visible         =   0   'False
      Width           =   5280
      Begin VSFlex8Ctl.VSFlexGrid vfgRecord 
         Height          =   1860
         Left            =   600
         TabIndex        =   14
         Top             =   240
         Width           =   4305
         _cx             =   7594
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
         BackColorFixed  =   14737632
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16635590
         ForeColorSel    =   -2147483640
         BackColorBkg    =   14737632
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   3
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
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
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   60
      Top             =   15
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmQCLJAverage.frx":7ACD
      Left            =   615
      Top             =   210
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmQCLJAverage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mColL  '��ֵ���ݱ���
    ��� = 0: ����: ���: ʵ������
End Enum

Const conPane_Record = 201
Const conPane_Charts = 202
Const conPane_Report = 203
Const conPane_Data = 204
'-----------------------------------------------------
'�������
'-----------------------------------------------------
Private mstrPrivs As String     '��ǰʹ����Ȩ�޴�
Private mlngListWidth As Long   '�б������ƿ��

Private mfrmChartLJAverage As frmQCChartLJAverage     'LJ����ͼ����

'-----------------------------------------------------
'��ʱ����
'-----------------------------------------------------
Dim cbrControl As CommandBarControl
Dim cbrCustom As CommandBarControlCustom
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar

Dim lngCount As Long
Private mstr��Ŀ As String                  '�洢�û���ǰѡ�е���Ŀ
Private mrsAverage As New ADODB.Recordset      '�����ֵ,SD����
Private mrsData As New ADODB.Recordset      '�����������ֵ����

Private mLastStartDate As Date, mLastEndDate As Date
Private mLastCell As String '�����뿪ǰ�ĵ�Ԫ��������������Ź���

Private Const ID_MENU_MOUSE = 90                                    '�Ҽ��˵�
Private mlngDeptID As Long
Private mlngMachineID As Long
Private mlngItemID As Long                                          '��ǰѡ�е���ĿID
Private mLastItemID As Long                                         '�ϴ���ʾ����ĿID�������ظ�ˢ��
'-----------------------------------------------------
'����Ϊ�ڲ���������
'-----------------------------------------------------
Private Function zlRefRecord() As Long
    '���ܣ�ˢ���ʿؽ����¼
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim lngCol As Long, lngRow As Long
    Dim dtStart As Date, dtEnd As Date

    Err = 0: On Error GoTo ErrHand
    If mlngItemID = 0 Then Exit Function
    
    dtStart = Format(Me.dtp����(0).Value, "yyyy-MM-dd")
    dtEnd = Format(Me.dtp����(1).Value, "yyyy-MM-dd 23:59:59")
    
    '��ȡָ��ʱ�䷶Χ��ͨ����˵ı걾 ������Ŀ�����ƽ��ֵ
    gstrSql = " Select Trunc(a.����ʱ��) As ����,Avg(Translate(Zl_To_Number(b.������,0),'>=<+-','00000')) As ��� " & _
                "From ����걾��¼ A, ������ͨ��� b " & _
                "Where a.����� Is Not Null And a.id=b.����걾ID And b.������Ŀid + 0 = [1] And a.����ʱ�� Between [2] And [3] " & _
                "Group By Trunc(a.����ʱ��)  order by Trunc(a.����ʱ��) "
            'Nvl(���ý��, 0) * -1 +
    Set mrsData = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngItemID, dtStart, dtEnd)
 
    '���ݻ��棬���ʿ�ͼʱ���õ�(frmQCChartLJAverage)
     
    With Me.vfgRecord
        .Redraw = flexRDNone
        .Clear
        .FixedCols = 3
        .Cols = .FixedCols
        .ExtendLastCol = False '���Զ���չ���һ�еĿ��
        .Rows = 4
        
        .ColWidth(0) = 1200
        .TextMatrix(mColL.���, 0) = ""
        .TextMatrix(mColL.���, 1) = "��ֵ": .ColWidth(1) = 500
        .TextMatrix(mColL.���, 2) = "SD": .ColWidth(2) = 500
        .TextMatrix(mColL.����, 0) = "����"
        .TextMatrix(mColL.���, 0) = "���"
        .TextMatrix(mColL.ʵ������, 0) = "ʵ������": .RowHidden(mColL.ʵ������) = True
        .ColAlignment(0) = flexAlignLeftCenter
        
        '����������ֵ��䵽����б���
        Do Until mrsData.EOF
            .Cols = .Cols + 1
            .TextMatrix(mColL.���, mrsData.AbsolutePosition + 2) = mrsData.AbsolutePosition
            .TextMatrix(mColL.����, mrsData.AbsolutePosition + 2) = Format(Nvl(mrsData!����), "yy-MM-dd")
            .TextMatrix(mColL.���, mrsData.AbsolutePosition + 2) = Round(Nvl(mrsData!���, 0), 2)
            .TextMatrix(mColL.ʵ������, mrsData.AbsolutePage + 2) = Nvl(mrsData!����)
            mrsData.MoveNext
        Loop
        
        '��д��ֵ��SD��CV
        'translate(b.������,'>=<-+','00000')
        gstrSql = "Select Round(Avg(���), 2) As ��ֵ, Round(Stddev(���), 3) As Sd " & _
                  "From (Select Trunc(a.����ʱ��) As ����,Avg(Translate(Zl_To_Number(b.������,0),'>=<+-','00000')) As ��� " & _
                        "From ����걾��¼ A, ������ͨ��� b Where a.����� Is Not Null And a.id=b.����걾ID " & _
                        "And b.������Ŀid + 0 = [1] And a.����ʱ�� Between [2] And [3] " & _
                  "Group By Trunc(a.����ʱ��))"
        Set mrsAverage = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngItemID, dtStart, dtEnd)
        
        '���ݻ��棬���ʿ�ͼʱ���õ�(frmQCChartLJAverage)
        
        If Not mrsAverage.EOF Then
            .TextMatrix(mColL.���, 1) = Val("" & mrsAverage!��ֵ)
            .TextMatrix(mColL.���, 2) = Val("" & mrsAverage!SD)
        End If
 
        If .Cols > .FixedCols Then
            .Cell(flexcpAlignment, mColL.���, .FixedCols, mColL.����, .Cols - 1) = flexAlignCenterCenter
            .AutoSize 0, .Cols - 1
        End If
        .Redraw = flexRDDirect
        If .Cols > .FixedCols Then .COL = .FixedCols
    End With
    
    zlRefRecord = Me.vfgRecord.Cols - Me.vfgRecord.FixedCols
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    zlRefRecord = 0
End Function

Private Sub zlRefOthers()
    '���ܣ�������ʾ���ԣ�ˢ�³��ʿؼ�¼��ͼ�κͱ���
    Dim strLists As String, intLists As Integer
    Dim lngItemID As Long, strFromDate As String, strToDate As String
    Dim str�������� As String

    If mlngItemID = 0 Then Exit Sub

    If mlngItemID = mLastItemID Then Exit Sub
    If mlngItemID = -1 Then mlngItemID = mLastItemID

    mLastItemID = mlngItemID
    lngItemID = mlngItemID
    strFromDate = Format(Me.dtp����(0).Value, "yyyy-MM-dd")
    strToDate = Format(Me.dtp����(1).Value, "yyyy-MM-dd")
    str�������� = Trim(Left$(Me.cbo����.Text, 30))
    
    '��õ�ǰѡ����ʿ�ͼ
    Dim intSelTab As Integer
    For lngCount = 0 To Me.tbcCharts.ItemCount - 1
        If Me.tbcCharts.Item(lngCount).Selected Then intSelTab = lngCount: Exit For
    Next

    If Me.tbcCharts.Item(intSelTab).Visible = False Then Me.tbcCharts.Item(0).Selected = True
    If Me.tbcCharts.Item(0).Selected Then Call mfrmChartLJAverage.zlRefresh(lngItemID, str��������, strFromDate, strToDate, mrsData, mrsAverage)
End Sub

Private Sub cbo����_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim lngMachineID As Long                '����ID
    
    On Error GoTo errH
    
    lngMachineID = mlngMachineID
    
    If Me.cbo����.ListCount <= 0 Then Exit Sub
    
    If InStr(1, mstrPrivs, "���п���") > 0 Then
        gstrSql = " Select Distinct D.ID, D.����, D.����, D.�ʿ�ˮƽ�� From �������� D " & _
                    " Where  Nvl(D.΢����, 0) <> 1 and d.ʹ��С��id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(Me.cbo����.ItemData(Me.cbo����.ListIndex)))
    Else
        gstrSql = "Select Distinct D.ID, D.����, D.����, D.�ʿ�ˮƽ��" & vbNewLine & _
                    " From �������� D " & vbNewLine & _
                    " Where Nvl(D.΢����, 0) <> 1 And D.ʹ��С��id = [2] And" & vbNewLine & _
                    "      D.ID In (Select Distinct D.ID" & vbNewLine & _
                    "               From ����С���Ա A, ����С�� B, ����С������ C, �������� D" & vbNewLine & _
                    "               Where A.С��id = B.ID And B.ID = C.С��id��and ��Աid = [1] And C.����id = D.ID)"

        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(UserInfo.ID), CLng(Me.cbo����.ItemData(Me.cbo����.ListIndex)))
    End If
    
    With rsTemp
        Me.cbo����.Clear
        
        Do While Not .EOF
            Me.cbo����.AddItem !���� & Space(200) & !�ʿ�ˮƽ��
            Me.cbo����.ItemData(Me.cbo����.NewIndex) = !ID
            If !ID = lngMachineID Then
                Me.cbo����.ListIndex = Me.cbo����.NewIndex
            End If
            .MoveNext
        Loop
        If Me.cbo����.ListCount > 0 And cbo����.ListIndex = -1 Then
            Me.cbo����.ListIndex = 0
        End If
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'-----------------------------------------------------
'����Ϊ�ؼ��¼�����
'-----------------------------------------------------
Private Sub cbo����_Click()
    Dim lngItemID As Long   '��ĿID

    Dim rsTemp As New ADODB.Recordset
    
    mlngItemID = Val(zlDatabase.GetPara("��Ŀ", glngSys, 1209, 0))
    

    If Me.cbo����.ListIndex = -1 Then Exit Sub
    Me.cbo����.Tag = Right(Me.cbo����.Text, 1)
    
    Err = 0: On Error GoTo ErrHand
    '��ȡ������ص�����  ������Ŀ
    gstrSql = "Select Distinct b.ID, b.����, b.Ӣ����, b.������" & vbNewLine & _
                " From ����������Ŀ a, ����������Ŀ b, ������Ŀ c" & vbNewLine & _
                " Where a.��Ŀid = b.ID And a.��Ŀid = c.������Ŀid And c.������� = 1 And ����ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(Me.cbo����.ItemData(Me.cbo����.ListIndex)))
    
    If rsTemp.RecordCount <= 0 Then MsgBox "��δ�������������Ŀ���ã�", vbInformation, gstrSysName: vfgItem.Clear:  Exit Sub
    
    With Me.vfgItem
        .FixedRows = 1
        .SelectionMode = flexSelectionByRow
        
        Set .DataSource = rsTemp
        .ColWidth(0) = 0
        .ColWidth(1) = 500
        .ColWidth(2) = 600
        .ColWidth(3) = 600
        .ColHidden(0) = True
        .AutoSize 1, 2
        .ColWidth(1) = 20
    End With
    Call vfgItem_RowColChange
        
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim rsTmp As New ADODB.Recordset
    
    '------------------------------------
    Select Case Control.ID
    
    Case conMenu_File_PrintSet
        Select Case Me.tbcCharts.Selected.Index
        Case 0: ReportPrintSet gcnOracle, glngSys, "ZL1_INSIDE_1208_8", Me
        End Select
    Case conMenu_File_Preview
        If Not Me.vfgRecord.Cols > Me.vfgRecord.FixedCols Then Exit Sub
        Select Case Me.tbcCharts.Selected.Index
        Case 0: Call mfrmChartLJAverage.ChartPrint: Call PrintQC(False)
        End Select
    Case conMenu_File_Print
        If Not Me.vfgRecord.Cols > Me.vfgRecord.FixedCols Then Exit Sub
        Select Case Me.tbcCharts.Selected.Index
        Case 0: Call mfrmChartLJAverage.ChartPrint: Call PrintQC(True)
        End Select
    Case conMenu_Edit_Save
        If Not Me.vfgRecord.Cols > Me.vfgRecord.FixedCols Then Exit Sub
        Select Case Me.tbcCharts.Selected.Index
        Case 0: Call mfrmChartLJAverage.ChartSaveAs
        End Select
    Case conMenu_Edit_MarkMap
        Select Case Me.tbcCharts.Selected.Index
        Case 0: Call mfrmChartLJAverage.ChartCopy
        End Select
    Case conMenu_File_Exit: Unload Me
    
    Case conMenu_View_ToolBar_Button
        Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each cbrControl In Me.cbsThis(2).Controls
            cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
        Me.cbsThis.RecalcLayout
    Case conMenu_View_StatusBar
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_Refresh
        mLastItemID = 0
        Call RefreshData
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_Tool_Reference_1
        '��
        Call ItemMoveUpDown(1)
    Case conMenu_Tool_Reference_2
        '��
        Call ItemMoveUpDown(2)
    Case Else
        If Control.ID < conMenu_ReportPopup * 100# + 1 Or Control.ID > conMenu_ReportPopup * 100# + 99 Then Exit Sub
        Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me)
    End Select
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_BatPrint, conMenu_Edit_Save, conMenu_Edit_MarkMap
        Control.Enabled = ((Me.vfgRecord.Cols > Me.vfgRecord.FixedCols) And mrsAverage.RecordCount <> 0)
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    End Select
End Sub

Private Sub cmdˢ��_Click()
    mLastItemID = 0
    Call RefreshData
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_Record
        Item.Handle = Me.picRecord.hWnd
    Case conPane_Charts
        Item.Handle = Me.picCharts.hWnd
    Case conPane_Data
        Item.Handle = Me.picData.hWnd
    End Select
End Sub

Private Sub dkpMan_RClick(ByVal Pane As XtremeDockingPane.IPane)
    If Pane.ID = conPane_Data Then
        Me.picData.Visible = True
    End If
End Sub

Private Sub RefreshData()
    Dim objControl As CommandBarControl
    Dim intRow As Integer
    Dim rsTemp As New ADODB.Recordset

    If mlngItemID = 0 Then Exit Sub

    Err = 0: On Error GoTo ErrHand
    
    If Me.dtp����(1).Value < Me.dtp����(0).Value Then
        MsgBox "�������ڲ��ܴ��ڿ�ʼ���ڣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    mLastStartDate = Format(dtp����(0).Value, "yyyy-MM-dd")
    mLastEndDate = Format(dtp����(1).Value, "yyyy-MM-dd")

    'ˢ�½������
    Call zlRefRecord
    Call zlRefOthers
    Call picRecord_Resize
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim lngDeptID As Long  '����ID
    '-----------------------------------------------------
    mlngListWidth = Me.picRecord.Width
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, False)
    
    lngDeptID = mlngDeptID
    
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Me.cbsThis.Icons = zlCommFun.GetPubIcons
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
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ������ͼ")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ����ͼ(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "������ͼ(&S)...")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "���ƿ���ͼ(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): cbrControl.BeginGroup = True
    End With
    
    
    Set cbrControl = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlLabel, 0, "����")
    cbrControl.Flags = xtpFlagRightAlign
    Set cbrCustom = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlCustom, 0, "����")
    cbrCustom.Handle = Me.cbo����.hWnd: cbrCustom.Flags = xtpFlagRightAlign
    Set cbrControl = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlLabel, 0, "����")
    cbrControl.Flags = xtpFlagRightAlign
    Set cbrCustom = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlCustom, 0, "����")
    cbrCustom.Handle = Me.cbo����.hWnd: cbrCustom.Flags = xtpFlagRightAlign
    cbrMenuBar.Visible = False
    
    '�����
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("C"), conMenu_Edit_MarkMap
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    
        .Add 0, VK_UP, conMenu_Tool_Reference_1
        .Add 0, VK_DOWN, conMenu_Tool_Reference_2
    
    End With
    
    '���ò����ò˵�
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_Edit_MarkMap
      '  .AddHiddenCommand conMenu_View_Refresh
    End With
    
    '-----------------------------------------------------
    '����������
    Set cbrToolBar = Me.cbsThis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "���Ϊ")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '-----------------------------------------------------
    '����ͣ������
    Dim panThis As Pane, panChild As Pane, panSub As Pane
    
    With Me.dkpMan
        Set panThis = .CreatePane(conPane_Record, 200, 400, DockLeftOf, Nothing)
        panThis.Title = "��ֵ�����"
        panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        Set panThis = .CreatePane(conPane_Charts, 400, 500, DockRightOf, Nothing)
        panThis.Title = "��ֵLJ�ʿ�ͼ"
        panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        
        Set panChild = .CreatePane(conPane_Data, 400, 100, DockBottomOf, panThis)
        panChild.Title = "������"
        panChild.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable

        panChild.Select
        
        .SetCommandBars Me.cbsThis
        .Options.ThemedFloatingFrames = True
        .Options.HideClient = True
    End With

    '-----------------------------------------------------
    '���ñ�񸽼Ӵ���
    Dim tbiThis As TabControlItem
    Set mfrmChartLJAverage = New frmQCChartLJAverage

    With Me.tbcCharts
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        Set tbiThis = .InsertItem(0, mfrmChartLJAverage.Caption, mfrmChartLJAverage.hWnd, 0)
        
    End With
    
    '-----------------------------------------------------
    '����ָ�
    Call RestoreWinState(Me, App.ProductName)
    
    '-----------------------------------------------------
    'װ���������
    Dim rsTemp As New ADODB.Recordset
    
    Me.dtp����(1).Value = zlDatabase.Currentdate: Me.dtp����(0).Value = CDate(Format(Me.dtp����(1).Value, "yyyy-MM") & "-01")
    Err = 0: On Error GoTo ErrHand
    
    If InStr(1, mstrPrivs, "���п���") > 0 Then
        gstrSql = " Select Distinct b.Id, b.���� , b.���� As ���� From �������� a ,���ű� b Where a.ʹ��С��ID = b.ID "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName)
        
    Else

        gstrSql = "Select Distinct B.ID, B.����, B.���� As ����" & vbNewLine & _
                " From �������� A, ���ű� B " & vbNewLine & _
                " Where A.ʹ��С��id = B.ID And" & vbNewLine & _
                "      A.ʹ��С��id In (Select Distinct D.ʹ��С��id" & vbNewLine & _
                "                   From ����С���Ա A, ����С�� B, ����С������ C, �������� D" & vbNewLine & _
                "                   Where A.С��id = B.ID And B.ID = C.С��id��and ��Աid = [1] And C.����id = D.ID)"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName, UserInfo.ID)
    End If
    
    Me.cbo����.Clear
    Do Until rsTemp.EOF
        Me.cbo����.AddItem rsTemp("����") & "-" & rsTemp("����")
        Me.cbo����.ItemData(Me.cbo����.NewIndex) = rsTemp("Id")
        If rsTemp("ID") = lngDeptID Then
            Me.cbo����.ListIndex = Me.cbo����.NewIndex
        End If
        rsTemp.MoveNext
    Loop
    If Me.cbo����.ListCount = 0 Then MsgBox "��δ�������ʹ��С������ã�", vbInformation, gstrSysName: Unload Me: Exit Sub
    If cbo����.ListIndex = -1 Then
        Me.cbo����.ListIndex = 0
    End If
    If Me.cbo����.ListCount = 1 Then Me.cbo����.Enabled = False
    
    mLastStartDate = CDate(0)
    mLastEndDate = CDate(0)
    
    
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    Dim panThis As Pane
    If Me.WindowState = vbMinimized Then Exit Sub
    Set panThis = Me.dkpMan.FindPane(conPane_Record)
    panThis.MinTrackSize.SetSize mlngListWidth / Screen.TwipsPerPixelX, panThis.MinTrackSize.Height
    panThis.MaxTrackSize.SetSize mlngListWidth / Screen.TwipsPerPixelX, panThis.MaxTrackSize.Height
    Me.dkpMan.RecalcLayout
    Me.dkpMan.NormalizeSplitters
    panThis.MinTrackSize.SetSize mlngListWidth / Screen.TwipsPerPixelX, panThis.MinTrackSize.Height
    panThis.MaxTrackSize.SetSize Screen.Width / Screen.TwipsPerPixelX, panThis.MaxTrackSize.Height
    Me.dkpMan.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload mfrmChartLJAverage
    Set mfrmChartLJAverage = Nothing
    
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub picCharts_Resize()
    Err = 0: On Error Resume Next
    With Me.tbcCharts
        .Left = Me.picCharts.ScaleLeft: .Width = Me.picCharts.ScaleWidth - .Left
        .Top = Me.picCharts.ScaleTop: .Height = Me.picCharts.ScaleHeight - .Top
    End With
End Sub

Private Sub picData_Resize()
    Err = 0: On Error Resume Next
    '�����б�
    With Me.vfgRecord
        .Left = Me.picData.ScaleLeft: .Width = Me.picData.ScaleWidth - .Left
        .Top = Me.picData.ScaleTop
        .Height = Me.picData.ScaleHeight - .Top
    End With
End Sub

Private Sub picRecord_Resize()
    Err = 0: On Error Resume Next
    
    Me.cmdˢ��.Left = Me.picRecord.ScaleWidth - Me.cmdˢ��.Width - 15
    Me.dtp����(1).Width = Me.picRecord.ScaleWidth - Me.cmdˢ��.Width - 15 - Me.dtp����(1).Left - 15
    Me.dtp����(0).Width = Me.dtp����(1).Width

    '��Ŀ�б�
    With Me.vfgItem
        .Left = Me.picRecord.ScaleLeft: .Width = Me.picRecord.ScaleWidth - .Left
        .Height = Me.picRecord.ScaleHeight - .Top
    End With
    
End Sub

Private Sub tbcCharts_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    mlngItemID = -1 'ǿ��ˢ��
    If Me.Visible Then Call zlRefOthers
End Sub

Private Sub vfgItem_RowColChange()

    
    If mLastStartDate <> CDate(0) And mLastEndDate <> CDate(0) Then
        Me.dtp����(0) = mLastStartDate
        Me.dtp����(1) = mLastEndDate
    
    Else
        Me.dtp����(0) = CDate(Format(Now, "yyyy-MM-01"))
        Me.dtp����(1) = CDate(Format(Now, "yyyy-MM-dd"))
    End If
    With Me.vfgItem
        If .Row >= .FixedRows Then
            mstr��Ŀ = Trim(.TextMatrix(.Row, 3)) & "/" & Trim(.TextMatrix(.Row, 2))
            If mlngItemID <> Val(.TextMatrix(.Row, 0)) Then
                mlngItemID = Val(.TextMatrix(.Row, 0))
                
                Call RefreshData
            End If
        End If
    End With
    
End Sub

Private Sub vfgRecord_EnterCell()
    With vfgRecord
        mLastCell = .Row & "," & .COL
    End With
End Sub

Private Sub vfgRecord_LeaveCell()
    With vfgRecord
        mLastCell = .Row & "," & .COL
    End With
End Sub

Private Sub vfgRecord_RowColChange()
    With vfgRecord
        mLastCell = .Row & "," & .COL
    End With
End Sub

Private Sub PrintQC(blnPrintMode As Boolean)
    '��ӡ��Ԥ���ʿ�ͼ
    '����           intPrintMode =1 ��ӡ =2 Ԥ��
    '               intPrintType 0=LJ 1=FQ 2=ZS 3=YD 4=CS 5=MN
    
    Dim rsTmp As New ADODB.Recordset
    Dim strPrintType As String                  '��Ӧ�ĵ���
    Dim strQCID As String                       '�ʿ�ƷID���ܻ�����","�ָ��Ķ��ID
    Dim lngQCID As Long                         '�����ʿ�ƷID
    Dim lngItemID As String                     '��ĿID
    Dim lngMachine As Long                      '����ID
    Dim intLoop As Integer
    
    
    On Error GoTo errH
    
    strPrintType = "ZL1_INSIDE_1208_8"
    gstrSql = "Select b.w, b.h " & vbNewLine & _
                " From Zlreports a, Zlrptitems b" & vbNewLine & _
                " Where a.Id = b.����id And a.��� = [1] And b.���� = '��ֵLJͼ'"
                
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strPrintType)
    'û���ҵ�ʱ�˳�
    If rsTmp.EOF Then
        MsgBox "�ڵ��ݶ�����û�ж���<��ֵLJͼ>,���ڵ����ж���һ����Ϊ<��ֵLJͼ>��ͼ���!", vbQuestion, Me.Caption
        Exit Sub
    End If
    
    If Dir(App.path & "\QCLJAverage_Tmp") <> "" Then
        With Me.chtCopy
            .Load App.path & "\QCLJAverage_Tmp"
            Kill App.path & "\QCLJAverage_Tmp"
            .Width = Nvl(rsTmp("w"), 1280 * Screen.TwipsPerPixelX)
            .Height = Nvl(rsTmp("h"), 500 * Screen.TwipsPerPixelY)
            .Header.Text = ""
            .ChartLabels.RemoveAll
            .ChartArea.Location.Top = -5
            .ChartArea.Location.Height = .ChartArea.Location.Height + 15
    
            .SaveImageAsJpeg App.path & "\QC_LJAverage" & ".jpg", 1000, False, False, False
        End With
    End If
    
    '�õ���ĿID
    If mlngItemID = 0 Then Exit Sub
    lngItemID = mlngItemID
    lngMachine = CLng(Me.cbo����.ItemData(Me.cbo����.ListIndex))
    
    If Dir(App.path & "\QC_LJAverage.jpg") <> "" Then
        Call ReportOpen(gcnOracle, glngSys, strPrintType, Me, _
        "��λ=" & gstrUnitName, "��ʼʱ��=" & dtp����(0).Value, "����ʱ��=" & dtp����(1).Value, "����=" & Left(Trim(cbo����.Text), 30), "��Ŀ=" & mstr��Ŀ, _
        "�����ֵ=" & Val("" & mrsAverage!��ֵ), "SD=" & Val("" & mrsAverage!SD), "��ֵLJͼ=" & App.path & "\QC_LJAverage.jpg", _
        IIf(blnPrintMode, 2, 1))
    End If
    
    If Dir(App.path & "\QC*.jpg") <> "" Then Kill App.path & "\QC*.jpg"
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ItemMoveUpDown(ByVal intUpDown As Integer)
    '���¼�����
    On Error Resume Next
    With Me.vfgItem
        If intUpDown = 1 Then
            If .Row - 1 > .FixedRows Then .Select .Row - 1, .COL
        Else
            If .Row + 1 < .Rows Then .Select .Row + 1, .COL
        End If
    End With
End Sub

Public Sub ShowMe(frmParent As Object, ByVal strPrivs As String, ByVal lngDeptID As Long, ByVal lngMachineID As Long)
    mstrPrivs = strPrivs
    mlngDeptID = lngDeptID
    mlngMachineID = lngMachineID
    
    Me.Show vbModal, frmParent
End Sub


