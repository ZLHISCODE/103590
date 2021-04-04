VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCheckScrip 
   BackColor       =   &H80000005&
   Caption         =   "���̺�����������"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8820
   ControlBox      =   0   'False
   Icon            =   "frmCheckScrip.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frmCheckScrip.frx":628A
   ScaleHeight     =   5985
   ScaleWidth      =   8820
   Begin VB.CommandButton cmdExport 
      Caption         =   "��ѯ(&S)"
      Height          =   350
      Left            =   7560
      TabIndex        =   3
      Top             =   120
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvwReport 
      Height          =   5280
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   9313
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDropMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img32"
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      OLEDropMode     =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "������"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "���"
         Object.Width           =   2470
      EndProperty
   End
   Begin VB.Label lblInfor 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   1290
      TabIndex        =   1
      Top             =   3540
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ѯ���ڡ����˷��ü�¼���ĳ����������PLSQL�е���ʹ���µı�"
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
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   6405
   End
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   0
      Picture         =   "frmCheckScrip.frx":6783
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "frmCheckScrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event StatusTextUpdate(ByVal strMSG As String) 'Ҫ�����������״̬������

Public Sub RefreshList()
    Call LoadData
End Sub

Private Sub ShowStatusInfor(ByVal strMSG As String)
    RaiseEvent StatusTextUpdate(strMSG)
End Sub

Private Sub LoadData()
    Dim strSQL As String, i As Long, objItem As ListItem
    Dim rstmp As ADODB.Recordset
    
    Call ShowStatusInfor("����ִ�У����Ե�......")
    
    strSQL = "Select Distinct A.Owner, A.Type, A.Name" & vbNewLine & _
            "From All_Source A, Zlsystems B" & vbNewLine & _
            "Where A.Owner = B.������ And Instr(A.Text, '���˷��ü�¼') > 0 And Substr(Text, 1, 2) <> '--'" & vbNewLine & _
            "Order By A.Type, A.Name"
            
    On Error GoTo errH
    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
     
    lvwReport.ListItems.Clear
    If Not rstmp Is Nothing Then
        For i = 1 To rstmp.RecordCount
            Set objItem = lvwReport.ListItems.Add(, "_" & i, rstmp!Name)
            
            objItem.SubItems(1) = "" & rstmp!owner
            objItem.SubItems(2) = "" & rstmp!Type
            
            rstmp.MoveNext
        Next
        Call ShowStatusInfor("���ҵ�" & rstmp.RecordCount & "����¼��")
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdExport_Click()
    Call LoadData
End Sub

Private Sub Form_Load()
    
    lvwReport.ColumnHeaders(2).Position = 1 '������
    lvwReport.ColumnHeaders(3).Position = 2 '���
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim sngWidth As Long '��С���
    
    sngWidth = IIf(ScaleWidth < 5600, 5600, ScaleWidth)
    
    cmdExport.Left = Me.ScaleLeft + Me.ScaleWidth - 100 - cmdExport.Width
    lvwReport.Width = sngWidth - 100
    lvwReport.Height = IIf(ScaleHeight - lvwReport.Top < 0, 0, ScaleHeight - lvwReport.Top)
 End Sub
Private Sub Form_Unload(Cancel As Integer)
    'If picStatus.Visible Then Cancel = 1
End Sub
 
Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = True
End Function

Public Sub subPrint(bytMode As Byte)
    '----------------------------------------------------------------------------------------
    '--����:���д�ӡ,Ԥ���������EXCEL
    '--����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '----------------------------------------------------------------------------------------
    Dim objPrint As New zlPrintLvw
    objPrint.Title.Text = "���̺ͺ����б�"
    Set objPrint.Body.objdata = lvwReport
    objPrint.BelowAppItems.Add "��ӡ�ˣ�" & gstrUserName
    objPrint.BelowAppItems.Add "��ӡʱ�䣺" & Format(Now, "yyyy��MM��dd��")
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrViewLvw objPrint, 1
          Case 2
              zlPrintOrViewLvw objPrint, 2
          Case 3
              zlPrintOrViewLvw objPrint, 3
      End Select
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
End Sub
 
Private Sub lvwReport_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Dim blnDesc As Boolean
    
    If ColumnHeader.Tag = "1" Then
        blnDesc = True
        ColumnHeader.Tag = ""
    Else
        blnDesc = False
        ColumnHeader.Tag = "1"
    End If
    lvwReport.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvwReport.SortOrder = lvwDescending
    Else
        lvwReport.SortOrder = lvwAscending
    End If
    lvwReport.Sorted = True
    
    If Not lvwReport.SelectedItem Is Nothing Then lvwReport.SelectedItem.EnsureVisible
End Sub
