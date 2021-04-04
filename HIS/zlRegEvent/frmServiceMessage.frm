VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Begin VB.Form frmServiceMessage 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin XtremeReportControl.ReportControl rptMain 
      Height          =   5475
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      _Version        =   589884
      _ExtentX        =   11880
      _ExtentY        =   9657
      _StockProps     =   0
      BorderStyle     =   1
      PreviewMode     =   -1  'True
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   2190
      Top             =   5670
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceMessage.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceMessage.frx":059A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceMessage.frx":0B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceMessage.frx":10CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceMessage.frx":1668
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceMessage.frx":1C02
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmServiceMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Object
Public mdatBegin As Date, mdatEnd As Date
Public mstr�Ǽ��� As String, mstr��Ϣ���� As String
Public mblnShowRead As Boolean, mblnFilter As Boolean

Public Sub ShowMe(frmMain As Object)
    Set mfrmMain = frmMain
End Sub

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
        .SetImageList img16
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
            Set objCol = .Add(0, "", 20, False)
            objCol.Groupable = False
            objCol.Alignment = xtpAlignmentCenter
            Set objCol = .Add(1, "��Ϣ����", 60, True)
            objCol.Alignment = xtpAlignmentCenter
            rptMain.GroupsOrder.Add objCol
            objCol.Visible = False
            Set objCol = .Add(2, "����", 200, True)
            objCol.Alignment = xtpAlignmentLeft
            Set objCol = .Add(3, "֪ͨԭ��", 200, True)
            objCol.Alignment = xtpAlignmentLeft
            Set objCol = .Add(4, "�Ǽ���", 60, True)
            objCol.Alignment = xtpAlignmentCenter
            Set objCol = .Add(5, "�Ǽ�ʱ��", 100, True)
            objCol.Alignment = xtpAlignmentCenter
            Set objCol = .Add(6, "ID", 0, False)
            objCol.Alignment = xtpAlignmentCenter
            objCol.Visible = False
        End With
        
    End With
    Call LoadMessage(True)
End Sub

Public Sub LoadMessage(blnFirst As Boolean)
    Dim objCol As ReportColumn
    Dim objRecord As ReportRecord
    Dim ObjItem As ReportRecordItem
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim datNow As Date, strTemp As String
    datNow = zlDatabase.Currentdate
'    If blnFirst Then
'        strTemp = " And [1] Between a.��ʼʱ�� And a.��ֹʱ�� "
'    Else
'        If mblnShowRead = False Then
'            strTemp = " And [1] Between a.��ʼʱ�� And a.��ֹʱ�� "
'        End If
'    End If
    With rptMain
        .Records.DeleteAll
        strSQL = "Select ID, ֪ͨ����, ����, ����, ҽ������, ��Ŀ����, ֪ͨԭ��, �Ǽ���, �Ǽ�ʱ��, ��������, ��ʼʱ��, ��ֹʱ��, ����ҽ������, ͣ�￪ʼʱ��, ����ʱ��" & vbNewLine & _
                "From (Select a.Id, a.֪ͨ����, b.����, c.���� As ����, Nvl(d.����, b.ҽ������) As ҽ������, e.���� As ��Ŀ����, a.֪ͨԭ��, a.�Ǽ���, a.�Ǽ�ʱ��, f.���� As ��������," & vbNewLine & _
                "              a.��ʼʱ��, a.��ֹʱ��, Null As ����ҽ������, Null As ͣ�￪ʼʱ��, a.����ʱ��" & vbNewLine & _
                "       From ���˷�����Ϣ��¼ A, �ٴ������Դ B, ���ű� C, ��Ա�� D, �շ���ĿĿ¼ E, ������Ϣ F" & vbNewLine & _
                "       Where a.��Դid = b.Id And b.����id = c.Id And b.ҽ��id = d.Id(+) And" & vbNewLine & _
                "             b.��Ŀid = e.Id And a.����id = f.����id And a.֪ͨ���� = 3  " & IIf(blnFirst, " And Sysdate Between a.��ʼʱ�� And a.��ֹʱ�� ", " And (a.��ʼʱ�� Between [3] And [4] Or a.��ֹʱ�� Between [3] And [4]) ") & vbNewLine & _
                "       Union" & vbNewLine & _
                "       Select a.Id, a.֪ͨ����, g.����, c.���� As ����, Nvl(d.����, b.ҽ������) As ҽ������, e.���� As ��Ŀ����, a.֪ͨԭ��, a.�Ǽ���, a.�Ǽ�ʱ��, Null As ��������," & vbNewLine & _
                "              a.��ʼʱ��, a.��ֹʱ��, Nvl(Ds.����, b.����ҽ������) As ����ҽ������, b.ͣ�￪ʼʱ��," & vbNewLine & _
                "              Case" & vbNewLine & _
                "                When a.����ʱ�� < Sysdate - 9998 Then" & vbNewLine & _
                "                 Null" & vbNewLine & _
                "                Else" & vbNewLine & _
                "                 a.����ʱ��" & vbNewLine & _
                "              End As ����ʱ��" & vbNewLine & _
                "       From (Select Min(ID) As ID, ֪ͨ����, ��¼id, Min(�Ǽ���) As �Ǽ���, Min(�Ǽ�ʱ��) As �Ǽ�ʱ��, Min(����id) As ����id, Min(֪ͨԭ��) As ֪ͨԭ��," & vbNewLine & _
                "                     Min(Nvl(����ʱ��, Sysdate - 9999)) As ����ʱ��, ��ʼʱ��, ��ֹʱ��" & vbNewLine & _
                "              From ���˷�����Ϣ��¼" & vbNewLine & _
                "              Where ֪ͨ���� In (1, 2)" & IIf(blnFirst, "", " And �Ǽ�ʱ�� Between [3] And [4] ") & vbNewLine & _
                "              Group By ֪ͨ����, ��¼id, ��ʼʱ��, ��ֹʱ��) A, �ٴ������¼ B, ���ű� C, ��Ա�� D, ��Ա�� Ds, �շ���ĿĿ¼ E, �ٴ������Դ G" & vbNewLine & _
                "       Where a.��¼id = b.Id And g.����id = c.Id And b.��Դid = g.Id And b.ҽ��id = d.Id(+) And b.����ҽ��id = Ds.Id(+) And" & vbNewLine & _
                "             b.��Ŀid = e.Id And a.֪ͨ���� In (1, 2))" & vbNewLine & _
                "Where 1 = 1 "
        
        If blnFirst Then
            strSQL = strSQL & " And ����ʱ�� Is Null"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, datNow)
        Else
            If Not mblnShowRead Then strSQL = strSQL & " And ����ʱ�� Is Null"
            If mstr��Ϣ���� <> "" Then
                strSQL = strSQL & " And ֪ͨ���� In (Select Column_Value From Table(f_Str2list([2])))"
            End If
            If mstr�Ǽ��� <> "" Then strSQL = strSQL & " And �Ǽ��� = [5]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, datNow, mstr��Ϣ����, mdatBegin, mdatEnd, mstr�Ǽ���)
        End If
        Do While Not rsTemp.EOF
            Select Case Val(rsTemp!֪ͨ����)
                Case 1
                    Set objRecord = .Records.Add()
                    Set ObjItem = objRecord.AddItem("")
                    If Not IsNull(rsTemp!����ʱ��) Then
                        ObjItem.Icon = 1
                    Else
                        ObjItem.Icon = 0
                    End If
                    Set ObjItem = objRecord.AddItem("ҽ��ͣ��")
                    objRecord.AddItem Nvl(rsTemp!����) & "," & Nvl(rsTemp!ҽ������) & _
                                "(" & Nvl(rsTemp!��Ŀ����) & ")" & "����" & Nvl(rsTemp!ͣ�￪ʼʱ��) & "ͣ��"
                    
                    objRecord.AddItem Nvl(rsTemp!֪ͨԭ��)
                    objRecord.AddItem Nvl(rsTemp!�Ǽ���)
                    objRecord.AddItem Nvl(rsTemp!�Ǽ�ʱ��)
                    objRecord.AddItem Val(Nvl(rsTemp!ID))
                    
                    If Not IsNull(rsTemp!����ʱ��) Then
                        objRecord.Item(1).ForeColor = vbBlue
                        objRecord.Item(2).ForeColor = vbBlue
                        objRecord.Item(3).ForeColor = vbBlue
                        objRecord.Item(4).ForeColor = vbBlue
                        objRecord.Item(5).ForeColor = vbBlue
                    End If
                Case 2
                    Set objRecord = .Records.Add()
                    Set ObjItem = objRecord.AddItem("")
                    If Not IsNull(rsTemp!����ʱ��) Then
                        ObjItem.Icon = 3
                    Else
                        ObjItem.Icon = 2
                    End If
                    Set ObjItem = objRecord.AddItem("ҽ������")
                    objRecord.AddItem Nvl(rsTemp!����) & "," & Nvl(rsTemp!ҽ������) & _
                                "(" & Nvl(rsTemp!��Ŀ����) & ")" & "����" & Nvl(rsTemp!ͣ�￪ʼʱ��) & "ͣ��" & _
                                ",��" & Nvl(rsTemp!����ҽ������) & "����"
                    objRecord.AddItem Nvl(rsTemp!֪ͨԭ��)
                    objRecord.AddItem Nvl(rsTemp!�Ǽ���)
                    objRecord.AddItem Nvl(rsTemp!�Ǽ�ʱ��)
                    objRecord.AddItem Val(Nvl(rsTemp!ID))
                    If Not IsNull(rsTemp!����ʱ��) Then
                        objRecord.Item(1).ForeColor = vbBlue
                        objRecord.Item(2).ForeColor = vbBlue
                        objRecord.Item(3).ForeColor = vbBlue
                        objRecord.Item(4).ForeColor = vbBlue
                        objRecord.Item(5).ForeColor = vbBlue
                    End If
                Case 3
                    Set objRecord = .Records.Add()
                    Set ObjItem = objRecord.AddItem("")
                    If Not IsNull(rsTemp!����ʱ��) Then
                        ObjItem.Icon = 5
                    Else
                        ObjItem.Icon = 4
                    End If
                    Set ObjItem = objRecord.AddItem("ԤԼ�Ǽ�")
                    objRecord.AddItem Nvl(rsTemp!��������) & "��������" & Nvl(rsTemp!��ʼʱ��) & "��" & Nvl(rsTemp!��ֹʱ��) & "�临��"
                    objRecord.AddItem Nvl(rsTemp!֪ͨԭ��)
                    objRecord.AddItem Nvl(rsTemp!�Ǽ���)
                    objRecord.AddItem Nvl(rsTemp!�Ǽ�ʱ��)
                    objRecord.AddItem Val(Nvl(rsTemp!ID))
                    If Not IsNull(rsTemp!����ʱ��) Then
                        objRecord.Item(1).ForeColor = vbBlue
                        objRecord.Item(2).ForeColor = vbBlue
                        objRecord.Item(3).ForeColor = vbBlue
                        objRecord.Item(4).ForeColor = vbBlue
                        objRecord.Item(5).ForeColor = vbBlue
                    End If
            End Select
            rsTemp.MoveNext
        Loop
        .Populate
        If .Rows.Count <> 0 Then
            .Rows.Row(0).Selected = True
            Call rptMain_SelectionChanged
        Else
            If Not mfrmMain Is Nothing Then Call mfrmMain.NoData
        End If
    End With
End Sub

Private Sub Form_Resize()
    rptMain.Width = Me.ScaleWidth
    rptMain.Height = Me.ScaleHeight
End Sub

Private Sub rptMain_SelectionChanged()
    With rptMain
        If .SelectedRows.Count = 0 Then Exit Sub
        If mfrmMain Is Nothing Then Exit Sub
        If .SelectedRows.Row(0).GroupRow = True Then Call mfrmMain.NoData: Exit Sub
        Call mfrmMain.LoadData(.SelectedRows.Row(0).Record.Item(6).Value)
    End With
End Sub
