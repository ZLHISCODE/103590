VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSendBillPrint 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���Ƶ��ݴ�ӡ"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   Icon            =   "frmSendBillPrint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Ԥ��(&V)"
      Height          =   350
      Left            =   3540
      TabIndex        =   4
      ToolTipText     =   "Ԥ����ǰ����"
      Top             =   4665
      Width           =   1100
   End
   Begin VB.CommandButton cmdSetup 
      Caption         =   "����(&S)"
      Height          =   350
      Left            =   2445
      TabIndex        =   3
      ToolTipText     =   "���õ�ǰ����"
      Top             =   4665
      Width           =   1100
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "��ӡ(&P)"
      Height          =   350
      Left            =   4800
      TabIndex        =   1
      ToolTipText     =   "��ӡ����ѡ��ĵ���"
      Top             =   4665
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "����(&X)"
      Height          =   350
      Left            =   5895
      TabIndex        =   2
      Top             =   4665
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvwBill 
      Height          =   3795
      Left            =   75
      TabIndex        =   0
      Top             =   750
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   6694
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "���ݺ�"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "���Ƶ���"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "˵��"
         Object.Width           =   6350
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   270
      Picture         =   "frmSendBillPrint.frx":058A
      Top             =   165
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSendBillPrint.frx":0E54
      Height          =   525
      Left            =   930
      TabIndex        =   5
      Top             =   120
      Width           =   6090
   End
End
Attribute VB_Name = "frmSendBillPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlng���ͺ� As Long
Private mint���� As Integer
Private mlngǰ��ID As Long
Private mint��ӡ��ʽ As Integer
Private mblnItem As Boolean

Public Sub ShowMe(ByVal lng���ͺ� As Long, ByVal int���� As Integer, frmParent As Object, Optional ByVal lngǰ��ID As Long)
'������lng���ͺ�=���η��͵ķ��ͺ�
'      int����=1-����,2-סԺ(���ݳ���,���ǵ��ó���)
    mlng���ͺ� = lng���ͺ�
    mint���� = int����
    mlngǰ��ID = lngǰ��ID
    
    On Error Resume Next
    Me.Show 1, frmParent
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()
'���ܣ������Ƶ��ݶ�Ӧ���Զ��屨�����Ԥ��
    If lvwBill.SelectedItem Is Nothing Then Exit Sub
    With lvwBill.SelectedItem
        Call ReportOpen(gcnOracle, glngSys, .Tag, Me, "NO=" & .Text, "����=" & Val(.ListSubItems(1).Tag), 1)
    End With
End Sub

Private Sub cmdPrint_Click()
'���ܣ���ѡ������Ƶ��ݽ��д�ӡ
    Dim i As Long, j As Long
    Dim blnALL As Boolean
    
    If lvwBill.SelectedItem Is Nothing Then Exit Sub
    For i = 1 To lvwBill.ListItems.Count
        If lvwBill.ListItems(i).Checked Then j = j + 1
    Next
    If j = 0 Then
        MsgBox "����ѡ����Ҫ��ӡ�����Ƶ��ݡ�", vbInformation, gstrSysName
        Exit Sub
    ElseIf j = lvwBill.ListItems.Count Then
        blnALL = True
    End If
    
    cmdPrint.Enabled = False
    Screen.MousePointer = 11
    For i = 1 To lvwBill.ListItems.Count
        With lvwBill.ListItems(i)
            If .Checked Then
                .Selected = True: .EnsureVisible: Me.Refresh
                Call ReportOpen(gcnOracle, glngSys, .Tag, Me, "NO=" & .Text, "����=" & Val(.ListSubItems(1).Tag), 2)
                
                '�Ѵ�ӡ������ɫ��ʶ
                .Checked = False: .ForeColor = vbBlue
                For j = 1 To .ListSubItems.Count
                    .ListSubItems(j).ForeColor = vbBlue
                Next
            End If
        End With
    Next
    Screen.MousePointer = 0
    cmdPrint.Enabled = True
    
    '�ֹ���ӡʱ��ȫ����ӡ��Ϻ��Զ��˳�
    If mint��ӡ��ʽ = 1 And blnALL Then
        Unload Me: Exit Sub
    ElseIf Visible Then
        cmdCancel.SetFocus
    End If
End Sub

Private Sub cmdSetup_Click()
'���ܣ������Ƶ��ݶ�Ӧ���Զ��屨���������
    If lvwBill.SelectedItem Is Nothing Then Exit Sub
    Call ReportPrintSet(gcnOracle, glngSys, lvwBill.SelectedItem.Tag, Me)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    mblnItem = False
    
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        For i = 1 To lvwBill.ListItems.Count
            lvwBill.ListItems(i).Checked = True
        Next
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        For i = 1 To lvwBill.ListItems.Count
            lvwBill.ListItems(i).Checked = False
        Next
    End If
End Sub

Private Sub Form_Load()
    '���Ƶ��ݴ�ӡ��ʽ:0-����ӡ,1-�ֹ���ӡ,2-�Զ���ӡ
    If mint���� = 1 And mlngǰ��ID = 0 Then
        mint��ӡ��ʽ = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "���﷢�͵��ݴ�ӡ", 1))
    Else
        mint��ӡ��ʽ = 1
    End If
    If mint��ӡ��ʽ = 0 Then Unload Me: Exit Sub
    
    Call RestoreListViewState(lvwBill, App.ProductName & "\" & Me.Name, "")
    If Not LoadBill Then Unload Me: Exit Sub
    If lvwBill.ListItems.Count = 0 Then Unload Me: Exit Sub
    mblnItem = False
    
    '�Զ���ӡ���˳�
    If mint��ӡ��ʽ = 2 Then
        Call cmdPrint_Click
        Unload Me: Exit Sub
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveListViewState(lvwBill, App.ProductName & "\" & Me.Name, "")
End Sub

Private Sub lvwBill_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvwBill, ColumnHeader.Index)
End Sub

Private Function LoadBill() As Boolean
'���ܣ���ȡ���η��Ϳ��Դ�ӡ�����Ƶ����б�
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objItem As ListItem
    
    lvwBill.ListItems.Clear
    
    On Error GoTo errH
    
    '���������˵�/���¼�����Ƶ���,���ݵ��ݱ�ŵ��ñ���(�൱��֪ͨ��)
    strSQL = "Select Distinct D.ID,A.NO,A.��¼����,D.���,D.����,D.˵��" & _
        " From ����ҽ������ A,����ҽ����¼ B,���Ƶ���Ӧ�� C,�����ļ�Ŀ¼ D" & _
        " Where A.���ͺ�=[1] And A.ҽ��ID=B.ID" & _
        " And B.������ĿID=C.������ĿID And C.Ӧ�ó���=[2]" & _
        " And C.�����ļ�ID=D.ID And D.����=5" & _
        " And D.ǰ�� IN([2],3) And D.��д IN(1,2)" & _
        " Order by A.NO"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng���ͺ�, mint����)
    For i = 1 To rsTmp.RecordCount
        Set objItem = lvwBill.ListItems.Add(, "_" & rsTmp!ID & "_" & rsTmp!NO & "_" & rsTmp!��¼����, rsTmp!NO)
        objItem.SubItems(1) = Nvl(rsTmp!����)
        objItem.SubItems(2) = Nvl(rsTmp!˵��)
        objItem.Tag = "ZLCISBILL" & Format(rsTmp!���, "00000") & "-1" '��Ӧ���Զ��屨����
        objItem.ListSubItems(1).Tag = rsTmp!��¼����
        objItem.Checked = True
        rsTmp.MoveNext
    Next
    LoadBill = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub lvwBill_DblClick()
    If mblnItem Then Call lvwBill_KeyPress(13)
End Sub

Private Sub lvwBill_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If mblnItem Then
        Item.Selected = True
        Item.EnsureVisible
    End If
End Sub

Private Sub lvwBill_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mblnItem = True
End Sub

Private Sub lvwBill_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call cmdSetup_Click
    End If
End Sub

Private Sub lvwBill_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnItem = False
End Sub
