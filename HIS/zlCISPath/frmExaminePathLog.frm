VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "CO70B6~1.OCX"
Begin VB.Form frmExaminePathLog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "·���������"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8070
   Icon            =   "frmExaminePathLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   8070
   StartUpPosition =   1  '����������
   Begin XtremeReportControl.ReportControl rptLog 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      _Version        =   589884
      _ExtentX        =   5318
      _ExtentY        =   1931
      _StockProps     =   0
      BorderStyle     =   2
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6120
      TabIndex        =   1
      Top             =   4920
      Width           =   1100
   End
End
Attribute VB_Name = "frmExaminePathLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngPathID As Long
Private mlngVersion As Long

Private Enum COL_LIST_LOG
    LOG_���� = 0
    LOG_����˵��
    LOG_������Ա
    LOG_����ʱ��
End Enum

Public Sub ShowMe(ByRef objFrmMain As Object, ByVal lngPathID As Long, ByVal lngVersion As Long)
    mlngPathID = lngPathID
    mlngVersion = lngVersion
    Me.Show 1, objFrmMain
End Sub

Private Sub InitReportColumnLog()
    Dim objCol As ReportColumn

    With rptLog
        '����˳�������(�������Ϊ����)�ı��,Ҫ��Find(�к�)��ItemIndex������,���Կ���Record(�к�)����������
        Set objCol = .Columns.Add(LOG_����, "��˽��", 100, True)
        Set objCol = .Columns.Add(LOG_����˵��, "���˵��", 200, True)
        Set objCol = .Columns.Add(LOG_������Ա, "�����", 80, True)
        Set objCol = .Columns.Add(LOG_����ʱ��, "���ʱ��", 140, True)

        For Each objCol In .Columns
            objCol.Editable = False
        Next

        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .TreeIndent = 0 '�з�����ʱ�������߱��ϻ�����һ������
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ���������..."
        End With

        .AutoColumnSizing = False
        .AllowColumnRemove = False
        .ShowGroupBox = False
        .ShowItemsInGroups = False
        .PreviewMode = True
        .MultipleSelection = False    '������SelectionChanged�¼�
'        .SetImageList Me.img16
        
'        .GroupsOrder.Add .Columns(LOG_����)
'        .GroupsOrder(0).SortAscending = True    '����֮��,��������в���ʾ,�����е������ǲ����
'
'        '����֮�����ʧȥ��¼���е�˳��,���ǿ�м���������
'        .SortOrder.Add .Columns(LOG_����)
'        .SortOrder(0).SortAscending = True
        .SortOrder.Add .Columns(LOG_����ʱ��)
        .SortOrder(0).SortAscending = False
    End With
End Sub

Private Sub LoadAduit(ByVal lngPathID As Long, ByVal lngVersion As Long)

    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim objRecord   As ReportRecord
    Dim objItem     As ReportRecordItem
    If lngPathID = 0 Then
        rptLog.Records.DeleteAll
        rptLog.Populate
        Exit Sub
    End If

    If gbln˫��� Then
        strSql = "Select Decode(����״̬, 1, 'ҽ������ͨ��', 2, 'ҽ������δ��', 3, 'ҩ�������ͨ��', 4, 'ҩ�������δ��') As ����״̬, NVL(����˵��,'δ��д') As ����˵��, ������Ա, ����ʱ��" & vbNewLine & _
            "From �ٴ�·�����" & vbNewLine & _
            "Where ·��id = [1] And �汾�� = [2]" & vbNewLine & _
            "Order By ����ʱ�� Desc"
    Else
        strSql = "Select Decode(����״̬, 1, '���ͨ��', 2, '���δ��', 3, 'ҩ�������ͨ��', 4, 'ҩ�������δ��') As ����״̬, NVL(����˵��,'δ��д') As ����˵��, ������Ա, ����ʱ��" & vbNewLine & _
            "From �ٴ�·�����" & vbNewLine & _
            "Where ·��id = [1] And �汾�� = [2]" & vbNewLine & _
            "Order By ����ʱ�� Desc"
    End If
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, lngVersion)
    
    rptLog.Records.DeleteAll
    Do While Not rsTmp.EOF
        Set objRecord = Me.rptLog.Records.Add()
        Set objItem = objRecord.AddItem(rsTmp!����״̬ & "")
        Set objItem = objRecord.AddItem(rsTmp!����˵�� & "")
        Set objItem = objRecord.AddItem(rsTmp!������Ա & "")
        Set objItem = objRecord.AddItem(Format(rsTmp!����ʱ�� & "", "YYYY-MM-DD HH:MM:SS"))
        rsTmp.MoveNext
    Loop

    rptLog.Populate
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call InitReportColumnLog
    Call LoadAduit(mlngPathID, mlngVersion)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.rptLog.Move 60, 60, Me.ScaleWidth - 120, Me.ScaleHeight - 700
    cmdOK.Move Me.ScaleWidth - cmdOK.Width - 240, Me.ScaleHeight - cmdOK.Height - 120
End Sub
