VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPageMedRecNOSel 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5595
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lvwInStationNo 
      Height          =   1995
      Left            =   300
      TabIndex        =   0
      Top             =   450
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   3519
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "סԺ��"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "����"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "��Ժ����"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "��Ժ����"
         Object.Width           =   2646
      EndProperty
   End
End
Attribute VB_Name = "frmPageMedRecNOSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mstrID As String                 'סԺ��_��ҳid
Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        mstrID = 0
        Unload Me
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        mstrID = Me.lvwInStationNo.SelectedItem.Key
        Unload Me
    End If
End Sub

Private Sub Form_Resize()
    Me.lvwInStationNo.Top = 0
    Me.lvwInStationNo.Left = 0
    Me.lvwInStationNo.Width = Me.ScaleWidth
    Me.lvwInStationNo.Height = Me.ScaleHeight
End Sub
Public Function ShowMe(objfrm As Object, rsPatient As ADODB.Recordset, strIf As String) As String
    '---------------------------------------------------------------------------------------------'
    '����        �ṩ���ϼ��������
    '����        objfrm�ϼ��������
    '            rsPatient���ݼ�
    '            Visible�Ƿ���ʾ��رմ���
    '---------------------------------------------------------------------------------------------'
    On Error GoTo errH
    
    If Me.Visible = True Then
        Unload Me
        ShowMe = "_"
        Exit Function
    End If
    
    Me.lvwInStationNo.ListItems.Clear
    If rsPatient.State = 1 Then
        rsPatient.Filter = strIf
        rsPatient.Sort = "סԺ���� Desc"
        If rsPatient.EOF = False Then
            rsPatient.MoveFirst
        End If
    ElseIf rsPatient.State = adStateClosed Then
        Unload Me
        ShowMe = "_"
        Exit Function
    End If
    
    Do Until rsPatient.EOF
        If Check�Ƿ���ڲ���(rsPatient("סԺ��"), rsPatient("סԺ����")) Then
            Set objList = Me.lvwInStationNo.ListItems.Add(, rsPatient("סԺ��") & "_" & rsPatient("סԺ����"), rsPatient("סԺ��"))
            objList.SubItems(1) = rsPatient("����")
            objList.SubItems(2) = Format(rsPatient("��Ժ����"), "yyyy-mm-dd")
            objList.SubItems(3) = Format(rsPatient("��Ժ����"), "yyyy-mm-dd")
        End If
        rsPatient.MoveNext
    Loop
    
    If Me.lvwInStationNo.ListItems.Count > 0 Then
        Me.Show vbModal, objfrm
        ShowMe = mstrID
    Else
        MsgBox "û���ҵ�������Ϣ!", vbInformation, gstrSysName
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub lvwInStationNo_Click()
    mstrID = Me.lvwInStationNo.SelectedItem.Key
End Sub

Private Sub lvwInStationNo_DblClick()
    mstrID = Me.lvwInStationNo.SelectedItem.Key
    Unload Me
End Sub
Private Function Check�Ƿ���ڲ���(strסԺ�� As String, lng��ҳID As Integer) As Boolean
    '----------------------------------------------------------------------------------------------------------
    '����:����Ƿ��ѽ�������
    '����:  strסԺ��-סԺ��
    '       lng��ҳID-��ҳID
    '����:True������ False����
    '----------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    gstrSQL = "" & _
        "   Select 1 from ������Ϣ a , ������ҳ b " & _
        "   Where a.����ID = b.����ID and b.��Ŀ���� is not null and b.��Ժ���� is not null  and " & _
        "         a.סԺ�� = [1] and b.��ҳID = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strסԺ��, lng��ҳID)
    If rsTmp.EOF Then
        Check�Ƿ���ڲ��� = True
    End If
    rsTmp.Close
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

