VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Begin VB.Form frmLabMainPrintFormat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ӡ���ݸ�ʽ����"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4845
   Icon            =   "frmLabMainPrintFormat.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin XtremeReportControl.ReportControl rptAlist 
      Height          =   2535
      Left            =   30
      TabIndex        =   1
      Top             =   300
      Width           =   4785
      _Version        =   589884
      _ExtentX        =   8440
      _ExtentY        =   4471
      _StockProps     =   0
      BorderStyle     =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��ѡ��һ�����ݸ�ʽ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1890
   End
End
Attribute VB_Name = "frmLabMainPrintFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mstrҽ��ID As String
Dim mstrPrintCode As String
Private Enum mCol
    ������
    ����
End Enum
Public Sub ShowMe(Objfrm As Object, strҽ�� As String, strPrintCode As String)
    Dim rsTmp As New ADODB.Recordset
    Dim Record As ReportRecord                      '�б����ݼ�
    Dim Column As ReportColumn
    Dim intLoop As Integer
    
    rptAlist.AllowColumnRemove = False
    rptAlist.ShowItemsInGroups = False
    With rptAlist.PaintManager
        .ColumnStyle = xtpColumnShaded
        .GridLineColor = RGB(225, 225, 225)
        .NoGroupByText = "�϶��б��⵽����,�����з���..."
        .NoItemsText = "û�п���ʾ����Ŀ..."
        .VerticalGridStyle = xtpGridSolid
        .HideSelection = True
    End With
    With Me.rptAlist.Columns
        Set Column = .Add(mCol.������, "������", 120, True)
        Set Column = .Add(mCol.����, "����", 120, True)
    End With

    gstrSql = "Select /*+ rule */" & vbNewLine & _
        " Distinct 'ZLCISBILL' || Trim(To_Char(C.���, '00000')) || '-2' As ������,  A.��¼����,c.����" & vbNewLine & _
        " From ����ҽ������ A, �����ļ��б� C, ����ҽ����¼ D, ��������Ӧ�� E" & vbNewLine & _
        " Where E.�����ļ�id = C.ID And D.������Ŀid = E.������Ŀid And A.ҽ��id = D.ID And " & vbNewLine & _
        " E.Ӧ�ó��� = Decode(D.������Դ, 2, 2, 4, 4, 1) And" & vbNewLine & _
        " D.���id In (Select * From Table(Cast(f_Num2list([1]) As ZLTOOLS.t_Numlist)))"

    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strҽ��)
    If rsTmp.RecordCount <= 1 Then
        If rsTmp.EOF = False Then
            strPrintCode = Nvl(rsTmp("������"))
        End If
        Unload Me
        Exit Sub
    End If
    Do While Not rsTmp.EOF
        Set Record = Me.rptAlist.Records.Add
        For intLoop = 0 To Me.rptAlist.Columns.Count - 1
            Record.AddItem ""
        Next
        Record(mCol.������).Value = Nvl(rsTmp("������"))
        Record(mCol.����).Value = Nvl(rsTmp("����"))
        rsTmp.MoveNext
    Loop
    Me.rptAlist.Populate
    mstrҽ��ID = strҽ��
    Me.Show vbModal, Objfrm
    strPrintCode = mstrPrintCode
End Sub

Private Sub rptAlist_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    mstrPrintCode = Item.Record(mCol.������).Value
    Unload Me
End Sub
