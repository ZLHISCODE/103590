VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Begin VB.Form frmAppRuleImport 
   Caption         =   "�����������"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8340
   Icon            =   "frmAppRuleImport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   8340
   StartUpPosition =   1  '����������
   Begin XtremeReportControl.ReportControl rptRule 
      Height          =   3000
      Left            =   120
      TabIndex        =   5
      Top             =   2625
      Width           =   8115
      _Version        =   589884
      _ExtentX        =   14314
      _ExtentY        =   5292
      _StockProps     =   0
      BorderStyle     =   2
      AllowColumnSort =   0   'False
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
   End
   Begin VB.Frame fraLine 
      Height          =   15
      Left            =   -60
      TabIndex        =   6
      Top             =   420
      Width           =   8355
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "����(&I)"
      Default         =   -1  'True
      Height          =   350
      Left            =   6975
      TabIndex        =   2
      Top             =   555
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6975
      TabIndex        =   3
      Top             =   990
      Width           =   1245
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   1785
      Left            =   120
      TabIndex        =   1
      Top             =   540
      Width           =   6720
      _cx             =   11853
      _cy             =   3149
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
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
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
      Rows            =   3
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
   Begin VB.Label lb��� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������жϹ���:"
      Height          =   180
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   1530
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "ѡ����ʵĶ����������Ϊ��ǰ�������ʿع���ע��ÿ�����������ˮƽ�����Ǻϡ�"
      Height          =   180
      Left            =   480
      TabIndex        =   0
      Top             =   135
      Width           =   7020
   End
   Begin VB.Image imgNote 
      Height          =   240
      Left            =   120
      Picture         =   "frmAppRuleImport.frx":058A
      Top             =   105
      Width           =   240
   End
End
Attribute VB_Name = "frmAppRuleImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mColL
    ������ = 0: ˮƽ��: ������
End Enum
Private Enum mColR
    ID = 0: �ж�: ����: ����Χ: ��ˮƽ: ���ϴ���: ��������
End Enum

Private mlngDevId As Long           '����id
Private mintLevel As Integer        '����ˮƽ��
Private mblnStart As Boolean        '�Ƿ��Ѿ������˶���������Ŀ,�ڹ����б�ˢ��ʱִ��
Private mlngGroupID As Long         '����ID
Private mblnOK As Boolean

'-----------------------------------------------------
'��ʱ����
'-----------------------------------------------------
Dim rptCol As ReportColumn
Dim rptRcd As ReportRecord
Dim rptItem As ReportRecordItem
Dim rptRow As ReportRow

Dim lngCount As Long

'-----------------------------------------------------
'����Ϊ�ڲ���������
'-----------------------------------------------------
Public Function ShowMe(frmParent As Form, lngDevId As Long, lngGroupID As Long) As Boolean
    '���ܣ�ˢ��װ��ָ������
    Dim rsTemp As New ADODB.Recordset
    
    mlngDevId = lngDevId
    mlngGroupID = lngGroupID
    
    gstrSql = "Select Max(Decode(A.�ʿ�ˮƽ��, Null, 1, 0, 1, A.�ʿ�ˮƽ��)) As ˮƽ��, Count(R.����) As ������" & vbNewLine & _
            "From �������� A, (Select ���� ,����ID From ������������ Where ����ID=[1] And ��ĿID=[2]) R" & vbNewLine & _
            "Where A.ID = R.����id(+)  And '0' = R.����(+) And A.ID = [1]"
    Err = 0: On Error GoTo ErrHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngDevId, mlngGroupID)
    mintLevel = rsTemp!ˮƽ��
    mblnStart = IIf(rsTemp!������ = 0, False, True)
    
    gstrSql = "Select Distinct ������, ˮƽ��, '������ÿ���������ˮƽ��Ϊ' || ˮƽ�� || '������...' As ������" & vbNewLine & _
        "From �����ʿط���" & vbNewLine & _
        "Where ˮƽ�� <= [1]" & vbNewLine & _
        "Order By ������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mintLevel)
    Set Me.vfgList.DataSource = rsTemp
    Me.vfgList.ColWidth(mColL.ˮƽ��) = 0
    Me.vfgList.ColHidden(mColL.ˮƽ��) = True
    Call vfgList_RowColChange
    
    mblnOK = False
    Me.Show vbModal, frmParent
    ShowMe = mblnOK
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    ShowMe = False
End Function

Private Function zlRefRule(strName As String) As Long
    '���ܣ�ˢ��װ�뵱ǰ�����Ĺ���
    Dim rsTemp As New ADODB.Recordset
    Dim objParent As Object, lngChilds As Long
    
    gstrSql = "Select ��� As ID, �ϼ� As �ϼ�id, ����, Decode(����, 'Y', '����: ', 'N', '����: ', '') || �ж� As �ж�, ������ As ����," & vbNewLine & _
            "       Decode(����Χ, 1, '��ǰ��', '��' || ����Χ || '��') As ����Χ, Decode(��ˮƽ, 1, '��', '') As ��ˮƽ," & vbNewLine & _
            "       Decode(Y����, 0, '��һ��', '����') As ���ϴ���, Decode(N����, 0, '��һ��', '����') As ��������" & vbNewLine & _
            "From (Select * From �����ʿط��� Where ������ = [1]) R" & vbNewLine & _
            "Start With �ϼ� Is Null" & vbNewLine & _
            "Connect By Prior ��� = �ϼ�" & vbNewLine & _
            "Order By Level, Decode(����, '0', 0, '1', 1, 'Y', 2, 'N', 3, 1)"
    Err = 0: On Error GoTo ErrHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strName)
    Err = 0: On Error GoTo 0
    Me.rptRule.Records.DeleteAll
    Me.rptRule.Populate
    With rsTemp
        Do While Not .EOF
            If Val("" & !�ϼ�ID) = 0 Then
                Set rptRcd = Me.rptRule.Records.Add()
            Else
                Me.rptRule.Populate
                For Each rptRow In Me.rptRule.Rows
                    If Val(rptRow.Record(mColR.ID).Value) = Val("" & !�ϼ�ID) Then
                        Set rptRcd = rptRow.Record.Childs.Add()
                    End If
                Next
            End If
            rptRcd.AddItem CStr(!ID)
            rptRcd.AddItem CStr(!�ж�)
            rptRcd.AddItem CStr("" & !����)
            rptRcd.AddItem CStr("" & !����Χ)
            rptRcd.AddItem CStr("" & !��ˮƽ)
            rptRcd.AddItem CStr("" & !���ϴ���)
            rptRcd.AddItem CStr("" & !��������)
            rptRcd.Expanded = True
            .MoveNext
        Loop
    End With
    Me.rptRule.Populate
    zlRefRule = Me.rptRule.Records.Count
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefRule = Me.rptRule.Records.Count
End Function

'-----------------------------------------------------
'����Ϊ�ؼ��¼�����
'-----------------------------------------------------
Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdImport_Click()
    With Me.vfgList
        If .Row < .FixedRows Then MsgBox "δѡ�з�����", vbInformation, gstrSysName: Exit Sub
        If mblnStart Then
            If MsgBox("����ʱ��ɾ����ǰ���Ѿ������˶���򣡼�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        If Val(.TextMatrix(.Row, mColL.ˮƽ��)) < mintLevel Then
            If MsgBox("ѡ������ˮƽ��С�ڵ�ǰ�������ʿ�ˮƽ����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        gstrSql = "Zl_�����ʿط���_Apply('" & Trim(.TextMatrix(.Row, mColL.������)) & "'," & mlngDevId & "," & mlngGroupID & ")"
    End With
    
    Err = 0: On Error GoTo ErrHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    
    mblnOK = True
    Unload Me: Exit Sub
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    '-----------------------------------------------------
    '�����б������
    With Me.rptRule
        .AutoColumnSizing = True
        Set rptCol = .Columns.Add(mColR.ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mColR.�ж�, "�ж�����", 150, True): rptCol.Editable = False: rptCol.Groupable = False
        rptCol.TreeColumn = True
        Set rptCol = .Columns.Add(mColR.����, "�жϹ���", 82, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mColR.����Χ, "����Χ", 45, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mColR.��ˮƽ, "��ˮƽ", 45, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mColR.���ϴ���, "���ϴ���", 55, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mColR.��������, "��������", 55, False): rptCol.Editable = False: rptCol.Groupable = False
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ����Ŀ..."
            .VerticalGridStyle = xtpGridSolid
        End With
    End With
    '-----------------------------------------------------
    '����ָ�
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    Me.fraLine.Left = 0: Me.fraLine.Width = Me.ScaleWidth
    Me.cmdCancel.Left = Me.ScaleWidth - Me.cmdCancel.Width - 120
    Me.cmdImport.Left = Me.ScaleWidth - Me.cmdImport.Width - 120
    Me.vfgList.Width = Me.cmdCancel.Left - Me.vfgList.Left * 2
    Me.rptRule.Width = Me.ScaleWidth - Me.rptRule.Left * 2
    Me.rptRule.Height = Me.ScaleHeight - Me.rptRule.Top - Me.rptRule.Left
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub vfgList_RowColChange()
    Call zlRefRule(Me.vfgList.TextMatrix(Me.vfgList.Row, mColL.������))
End Sub
