VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Begin VB.Form frmLabAuditingCourse 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "�걾������־"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7200
   Icon            =   "frmLabAuditingCourse.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin XtremeReportControl.ReportControl rptList 
      Height          =   3675
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6705
      _Version        =   589884
      _ExtentX        =   11827
      _ExtentY        =   6482
      _StockProps     =   0
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   5760
      TabIndex        =   1
      Top             =   4140
      Width           =   1100
   End
End
Attribute VB_Name = "frmLabAuditingCourse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mlngKey As Long             '�걾ID
Private Enum mCol
    ID = 0: �걾ID: ��������: ����Ա: ����ʱ��
End Enum

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim Column As ReportColumn
    
    With Me.rptList.Columns
        
        rptList.AllowColumnRemove = False
        rptList.ShowItemsInGroups = False
        
        With rptList.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ����Ŀ..."
            .VerticalGridStyle = xtpGridSolid
        End With

        Set Column = .Add(mCol.ID, "ID", 65, True): Column.Visible = False
        Set Column = .Add(mCol.�걾ID, "�걾Id", 65, True): Column.Visible = False
        Set Column = .Add(mCol.��������, "��������", 85, True)
        Set Column = .Add(mCol.����Ա, "����Ա", 95, True)
        Set Column = .Add(mCol.����ʱ��, "����ʱ��", 150, True)
    End With
    
    Call RefreshData
End Sub

Private Sub Form_Resize()
    With Me.rptList
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.cmdOK.Top - 200
    End With
End Sub

Public Sub ShowMe(objfrm As Object, lngKey As Long)
    mlngKey = lngKey
    Me.Show vbModal, objfrm
End Sub

Private Sub RefreshData()
    Dim rsTmp As New ADODB.Recordset
    Dim intLoop As Integer
    Dim Record As ReportRecord
    
    On Error GoTo errH
    gstrSql = "select id,�걾ID,decode(��������,0,'���',1,'ȡ�����',2,'���ղ���',3,'�ع�',4,'�޸ı��浥') as ��������,����Ա,����ʱ�� " & vbNewLine & _
              " from ���������¼ where �걾id =[1] order by id"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngKey)
    
    Do Until rsTmp.EOF
        Set Record = Me.rptList.Records.Add
        For intLoop = 0 To Me.rptList.Columns.Count
            Record.AddItem ""
        Next
        Record(mCol.ID).Value = Nvl(rsTmp("ID"))
        Record(mCol.�걾ID).Value = Nvl(rsTmp("�걾Id"))
        Record(mCol.��������).Value = Nvl(rsTmp("��������"))
        Record(mCol.����Ա).Value = Nvl(rsTmp("����Ա"))
        Record(mCol.����ʱ��).Value = Nvl(rsTmp("����ʱ��"))
        rsTmp.MoveNext
    Loop
    Me.rptList.Populate
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
