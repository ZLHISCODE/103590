VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Begin VB.Form frmAppRequestMain 
   BorderStyle     =   0  'None
   Caption         =   "frmAppRequestMain"
   ClientHeight    =   5550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin XtremeReportControl.ReportControl rptMain 
      Height          =   3765
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   5535
      _Version        =   589884
      _ExtentX        =   9763
      _ExtentY        =   6641
      _StockProps     =   0
      BorderStyle     =   1
      PreviewMode     =   -1  'True
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
   End
End
Attribute VB_Name = "frmAppRequestMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mstr�Ǽ��� As String
Public mdat��ʼʱ�� As Date
Public mdat����ʱ�� As Date
Public mdat����ʼ As Date
Public mdat������� As Date
Public mstr������ As String
Public mbln��ʾ���� As Boolean
Public mbyt���﷽ʽ As Byte
Public mbln�Ǽ�ʱ�� As Boolean
Public mbln����ʱ�� As Boolean

Private Sub Form_Load()
    Dim objCol As ReportColumn
    Dim objRecord As ReportRecord
    Dim ObjItem As ReportRecordItem
    With rptMain
        .AutoColumnSizing = False '��ʹ���Զ��п�
        .AllowColumnRemove = False '�������϶�ɾ����
        .ShowGroupBox = True '��ʾ�����
        .ShowItemsInGroups = False '����ʾ�ѷ������
        .MultipleSelection = False '���������ѡ��
        .PreviewMode = False
        .AllowColumnReorder = False

        With .PaintManager
            .HighlightBackColor = 16772055
            .HighlightForeColor = RGB(0, 0, 0)
            .MaxPreviewLines = 1
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(180, 180, 180)
            .VerticalGridStyle = xtpGridSolid
            .HorizontalGridStyle = xtpGridSolid '�������߸�ʽ
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .GroupBoxBackColor = RGB(180, 180, 180)
            .NoItemsText = ""
        End With
        
        With .Columns
            Set objCol = .Add(0, "����", 40, True)
            objCol.Alignment = xtpAlignmentCenter
            Set objCol = .Add(1, "����", 60, True)
            objCol.Alignment = xtpAlignmentCenter
            Set objCol = .Add(2, "����", 100, True)
            objCol.Alignment = xtpAlignmentCenter
            Set objCol = .Add(3, "��Ŀ", 200, True)
            objCol.Alignment = xtpAlignmentLeft
            Set objCol = .Add(4, "���﷽ʽ", 100, True)
            objCol.Alignment = xtpAlignmentLeft
            rptMain.GroupsOrder.Add objCol
            objCol.Visible = False
            Set objCol = .Add(5, "����ԭ��", 200, True)
            objCol.Alignment = xtpAlignmentLeft
            Set objCol = .Add(6, "��ʼ����", 150, True)
            objCol.Alignment = xtpAlignmentLeft
            Set objCol = .Add(7, "��ֹ����", 150, True)
            objCol.Alignment = xtpAlignmentLeft
            Set objCol = .Add(8, "��������", 100, True)
            objCol.Alignment = xtpAlignmentLeft
            Set objCol = .Add(9, "�Ǽ���", 100, True)
            objCol.Alignment = xtpAlignmentLeft
            Set objCol = .Add(10, "�Ǽ�ʱ��", 200, True)
            objCol.Alignment = xtpAlignmentLeft
            Set objCol = .Add(11, "������", 100, True)
            objCol.Alignment = xtpAlignmentLeft
            Set objCol = .Add(12, "����ʱ��", 200, True)
            objCol.Alignment = xtpAlignmentLeft
        End With
    End With
    Call LoadRecord(True)
End Sub

Private Sub Form_Resize()
    With rptMain
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight
    End With
End Sub

Private Sub LoadRecord(Optional ByVal blnFirst As Boolean)
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim objCol As ReportColumn
    Dim objRecord As ReportRecord
    Dim ObjItem As ReportRecordItem
    strSQL = "Select a.id As ��Ϣid, a.����, d.����, c.���� As ��Ŀ, e.���� As ����, a.ҽ������ As �Һ�ҽ��, a.��ʼʱ��, a.��ֹʱ��, b.���� As ��������, b.�����, b.�Ա�, b.����, b.��ͥ�绰 As ��ϵ�绰, a.�Ǽ���," & vbNewLine & _
            "       a.�Ǽ�ʱ��, a.֪ͨԭ�� As �Ǽ�ԭ��, a.����ʱ��, a.������, a.����˵��, a.���﷽ʽ , a.���� " & vbNewLine & _
            "From ���˷�����Ϣ��¼ A, ������Ϣ B, �շ���ĿĿ¼ C, �ٴ������Դ D, ���ű� E " & vbNewLine & _
            "Where a.��Ŀid = c.Id And a.����id = b.����id And a.֪ͨ���� = 3 And a.��Դid = d.id And d.����id = e.id "
    If blnFirst Then
        strSQL = strSQL & "And a.�Ǽ��� = [1] And a.�Ǽ�ʱ�� Between Trunc(Sysdate) And Trunc(Sysdate+1)-1/24/60/60 And a.����ʱ�� Is Null"
        mstr�Ǽ��� = UserInfo.����
    Else
        If mstr�Ǽ��� <> "" Then strSQL = strSQL & " And a.�Ǽ���=[1]"
        If mbln�Ǽ�ʱ�� Then
            strSQL = strSQL & " And a.�Ǽ�ʱ�� Between [2] And [3]"
        End If
        If mbln��ʾ���� = False Then
            strSQL = strSQL & " And a.����ʱ�� Is Null"
        Else
            If mstr������ <> "" Then strSQL = strSQL & " And (a.������=[4] Or a.������ Is Null)"
            If mbln����ʱ�� Then
                strSQL = strSQL & " And (a.����ʱ�� Between [5] And [6] Or a.����ʱ�� Is Null)"
            End If
        End If
        If mbyt���﷽ʽ <> 0 Then strSQL = strSQL & " And a.���﷽ʽ = [7]"
    End If
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr�Ǽ���, mdat��ʼʱ��, mdat����ʱ��, mstr������, mdat����ʼ, mdat�������, mbyt���﷽ʽ)
    With rptMain
        .Records.DeleteAll
        Do While Not rsTemp.EOF
            Set objRecord = .Records.Add
            objRecord.Tag = Val(Nvl(rsTemp!��ϢID))
            Set ObjItem = objRecord.AddItem(Nvl(rsTemp!����))
            Set ObjItem = objRecord.AddItem(Nvl(rsTemp!����))
            Set ObjItem = objRecord.AddItem(Nvl(rsTemp!����))
            Set ObjItem = objRecord.AddItem(Nvl(rsTemp!��Ŀ))
            Select Case Val(rsTemp!���﷽ʽ)
            Case 1
                Set ObjItem = objRecord.AddItem(rsTemp!���� & "���Ƴ̺���")
            Case 2
                Set ObjItem = objRecord.AddItem(rsTemp!���� & "���º���")
            Case 3
                Set ObjItem = objRecord.AddItem(rsTemp!���� & "�ܺ���")
            Case 4
                Set ObjItem = objRecord.AddItem(rsTemp!���� & "�����")
            End Select
            Set ObjItem = objRecord.AddItem(Nvl(rsTemp!�Ǽ�ԭ��))
            Set ObjItem = objRecord.AddItem(Format(rsTemp!��ʼʱ��, "yyyy-mm-dd"))
            Set ObjItem = objRecord.AddItem(Format(rsTemp!��ֹʱ��, "yyyy-mm-dd"))
            Set ObjItem = objRecord.AddItem(Nvl(rsTemp!��������))
            Set ObjItem = objRecord.AddItem(Nvl(rsTemp!�Ǽ���))
            Set ObjItem = objRecord.AddItem(Format(rsTemp!�Ǽ�ʱ��, "yyyy-MM-dd hh:mm:ss"))
            Set ObjItem = objRecord.AddItem(Nvl(rsTemp!������))
            Set ObjItem = objRecord.AddItem(Nvl(rsTemp!����ʱ��))
            
            If Not IsNull(rsTemp!����ʱ��) Then
                For i = 0 To 12
                    objRecord.Item(i).ForeColor = vbBlue
                Next i
            End If
            rsTemp.MoveNext
        Loop
        .Populate
    End With
End Sub

Public Sub RefreshData()
    Call LoadRecord(False)
End Sub

Private Sub rptMain_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If rptMain.SelectedRows.Count = 0 Then Exit Sub
    If rptMain.SelectedRows.Row(0).Record Is Nothing Then Exit Sub
    Call frmAppRequestEdit.ReadBill(Me, Val(rptMain.SelectedRows.Row(0).Record.Tag))
End Sub
