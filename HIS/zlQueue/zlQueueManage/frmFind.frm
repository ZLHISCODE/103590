VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFind 
   Caption         =   "����"
   ClientHeight    =   5268
   ClientLeft      =   72
   ClientTop       =   360
   ClientWidth     =   8580
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5268
   ScaleWidth      =   8580
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   852
      ScaleWidth      =   8580
      TabIndex        =   7
      Top             =   4410
      Width           =   8580
      Begin VB.CommandButton cmdExit 
         Caption         =   "�˳�(&Q)"
         Height          =   375
         Index           =   0
         Left            =   7080
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "����(&C)"
         Height          =   375
         Index           =   1
         Left            =   4680
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "�ָ�(&R)"
         Height          =   375
         Index           =   2
         Left            =   5880
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      ScaleHeight     =   972
      ScaleWidth      =   8580
      TabIndex        =   1
      Top             =   0
      Width           =   8580
      Begin VB.ComboBox cboFindWay 
         Height          =   300
         ItemData        =   "frmFind.frx":000C
         Left            =   1080
         List            =   "frmFind.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   120
         Width           =   2655
      End
      Begin VB.TextBox txtFindData 
         Height          =   300
         Left            =   1080
         TabIndex        =   3
         Top             =   555
         Width           =   2655
      End
      Begin VB.CommandButton cmdStartFind 
         Caption         =   "��ʼ����(&F)"
         Height          =   375
         Left            =   3840
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label labFindWay 
         Caption         =   "���ҷ�ʽ��"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   165
         Width           =   975
      End
      Begin VB.Label labFindData 
         Caption         =   "�������ݣ�"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   975
      End
   End
   Begin MSComctlLib.ListView lvwQueueData 
      Height          =   3255
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   8055
      _ExtentX        =   14203
      _ExtentY        =   5736
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "��������"
         Text            =   "��������"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "�ŶӺ���"
         Text            =   "�ŶӺ���"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "��������"
         Text            =   "��������"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "��������"
         Text            =   "��������"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "ҽ������"
         Text            =   "ҽ������"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "��������"
         Text            =   "��������"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "�������"
         Text            =   "�������"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Key             =   "�Ŷ�ʱ��"
         Text            =   "�Ŷ�ʱ��"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Key             =   "�Ŷ�״̬"
         Text            =   "�Ŷ�״̬"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mcnOracle As ADODB.Connection  'oracle ���ݿ�����
Private mstrFindKey As String          '���ݵĲ��ҷ�ʽ
Private gbyt�ſ� As Byte               '�ſ��ų���

Public Sub ShowFind(cnOracle As ADODB.Connection, ByVal lngCardLen As Long, Optional owner As Form = Null)
    Set mcnOracle = cnOracle
    
    gbyt�ſ� = lngCardLen
    Me.Show 1, owner
End Sub



Private Sub cmdExit_Click(Index As Integer)
    Dim strQueueId As String
    
    On Error GoTo errHandle
    
    Select Case Index
        Case 0
            Unload Me
        Case 1, 2
            strQueueId = GetSelectId()
          
            If Trim(strQueueId) = "" Then
                MsgBox "��δѡ��һ����Ҫ���и�����������ݡ�", vbInformation, "�Ŷӽк�ϵͳ"
                Exit Sub
            End If
            
            If Index = 1 Then
                Call Execute_����(Val(strQueueId))
            ElseIf Index = 2 Then
                Call Execute_�ָ�(Val(strQueueId))
            End If
            
            'ˢ������
            Call cmdStartFind_Click
            
            'MsgBox "����ִ����ɡ�", vbInformation, "�Ŷӽк�ϵͳ"
    End Select
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Function GetSelectId() As String
'***************************************
'
'ȡ�õ�ǰѡ�е�����
'
'***************************************
    On Error GoTo errHandle
        
        If lvwQueueData.SelectedItem Is Nothing Then
          GetSelectId = ""
          Exit Function
        End If
        
        GetSelectId = lvwQueueData.SelectedItem.Tag
        
    Exit Function
errHandle:
      GetSelectId = ""
      If ErrCenter = 1 Then Resume
End Function


Private Sub cmdStartFind_Click()
    Dim rsData As ADODB.Recordset
    Dim strFindType As String
    Dim strFindValue As String
    
    On Error GoTo errHandle
    strFindValue = txtFindData.Text
    
    If Trim(strFindValue) = "" Then
        MsgBox "��������Ҫ���ҵ�����ֵ��", vbOKOnly, Me.Caption
        
        Call txtFindData.SetFocus
        Exit Sub
    End If
    
    Call lvwQueueData.ListItems.Clear
    
    'ȡ�ü�������
    strFindType = cboFindWay.Text
    
    Set rsData = FindQueueData(strFindType, strFindValue)
    
    If rsData Is Nothing Then
        MsgBox "û�м������������ݡ�", vbInformation, "�Ŷӽк�ϵͳ"
        Exit Sub
    End If
    
    If rsData.RecordCount <= 0 Then
        MsgBox "û�м������������ݡ�", vbInformation, "�Ŷӽк�ϵͳ"
        Exit Sub
    End If
    
    Call LoadDataToFace(lvwQueueData, rsData, "ID")
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub



Private Function FindQueueData(ByVal findType As String, ByVal findData As String) As ADODB.Recordset
    Dim strSql As String, strFilter As String
    Dim str����� As String, str���� As String, str���￨�� As String, strҽ���� As String, str�Һŵ��� As String
    
    On Error GoTo errHandle
    
    strFilter = ""
    
    Select Case findType  ' '0-�����;1-����;2-�Һŵ�;3-���￨��;4-ҽ����
    Case "�����"
        str����� = Val(findData)
        strFilter = strFilter & " And A.����� = [1]"
    Case "����"
        str���� = findData & "%"
        strFilter = strFilter & " And A.���� Like [2]"
    Case "���￨��"
        str���￨�� = findData
        strFilter = strFilter & " And A.���￨��=[3]"
    Case Else    ' "ҽ����"
        strҽ���� = findData
        strFilter = strFilter & " And A.ҽ����=[4]"
    End Select
    
            
    If Trim(findType) <> "����" Then
        strSql = "Select q.ID, q.��������, p.���� as ��������, q.��������, q.�ŶӺ���, q.���� as ��������, " & _
                 " q.ҽ������, q.�������, q.�Ŷ�ʱ��, decode(q.�Ŷ�״̬, 1, '������', 0, '�Ŷ���', 3, '��ͣ', 4, '���', '������') as �Ŷ�״̬  " & vbCrLf & _
                 " From ������Ϣ A, �ŶӽкŶ��� Q, ���ű� P " & vbCrLf & _
                 " Where Q.����id = A.����ID and Q.����ID=P.ID " & vbCrLf & strFilter
    Else
        strSql = "Select q.ID, q.��������, p.���� as ��������, q.��������, q.�ŶӺ���, q.���� as ��������, " & _
                 " q.ҽ������, q.�������, q.�Ŷ�ʱ��, decode(q.�Ŷ�״̬, 1, '������', 0, '�Ŷ���', 3, '��ͣ', 4, '���', '������') as �Ŷ�״̬  " & vbCrLf & _
                 " From �ŶӽкŶ��� Q, ���ű� P " & vbCrLf & _
                 " Where Q.����ID=P.ID and Q.�������� like [2]"
    End If

    Set FindQueueData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, str�����, str����, str���￨��, strҽ����)
    
    Exit Function
errHandle:
    Set FindQueueData = Nothing
    If ErrCenter = 1 Then Resume

End Function



Private Sub LoadDataToFace(lvwData As ListView, rsData As ADODB.Recordset, strKey As String)
'**************************************************************************************************
'�����ѯ��������Ϣ��ListView��
'
'lvwQueueData������������ʾ
'rsData������Դ
'strKey������ؼ���
'
'**************************************************************************************************
    
    On Error GoTo errHandle

    '�����������
    Call lvwData.ListItems.Clear
    
    If rsData.RecordCount <= 0 Then Exit Sub
      
    Dim i As Integer
      
    Call rsData.MoveFirst
      
        
    'ѭ����ȡ����
    While Not rsData.EOF
      Dim liRow As ListItem
      
      Set liRow = lvwData.ListItems.Add()
      
      'liRow.SmallIcon = 1
      'liRow.Icon = 1
      
      '��ʹ��RestoreWinState���̺�listview�ؼ���ǰ���Զ����"_"
      '��ȡ��һ����Ϣ
      If Not IsNull(rsData.Fields.Item(Replace(lvwData.ColumnHeaders(1).Key, "_", ""))) Then
        liRow.Text = rsData.Fields.Item(Replace(lvwData.ColumnHeaders(1).Key, "_", ""))
      Else
        liRow.Text = ""
      End If
      
      '��ȡ�ؼ���
      liRow.Tag = rsData(strKey)
      
      For i = 2 To lvwData.ColumnHeaders.Count
        Dim liSubItem As ListSubItem
        
        Set liSubItem = liRow.ListSubItems.Add()
        
        If Not IsNull(rsData.Fields.Item(Replace(lvwData.ColumnHeaders(i).Key, "_", ""))) Then
          liSubItem.Text = rsData.Fields.Item(Replace(lvwData.ColumnHeaders(i).Key, "_", ""))
        Else
          liSubItem.Text = ""
        End If
    
      Next i
          
      Call rsData.MoveNext
    Wend
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub Execute_�ָ�(ByVal Id As Long)
    On Error GoTo errHandle
        
        Dim strSql As String
        
        strSql = "ZL_�ŶӽкŶ���_�ָ�(" & Id & ")"
                
        Call zlDatabase.ExecuteProcedure(strSql, "����")
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub Execute_����(ByVal Id As Long)
    On Error GoTo errHandle
        
        Dim strSql As String
        
        strSql = "ZL_�ŶӽкŶ���_����(" & Id & ")"
                
        Call zlDatabase.ExecuteProcedure(strSql, "����")
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub Form_Load()
    '�ָ�����״̬
    Call RestoreWinState(Me, App.ProductName)

    cboFindWay.ListIndex = 1
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    
    lvwQueueData.Left = 100
    lvwQueueData.Top = Picture1.Height + 100
    lvwQueueData.Width = Me.Width - 200
    lvwQueueData.Height = Picture2.Top - Picture1.Height - 200
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub Picture2_Resize()
    On Error Resume Next
    
    cmdExit(0).Left = Me.Width - cmdExit(0).Width - 200
    cmdExit(2).Left = cmdExit(0).Left - cmdExit(2).Width - 50
    cmdExit(1).Left = cmdExit(2).Left - cmdExit(1).Width - 50
End Sub

Private Sub txtFindData_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean
    Dim rsData As ADODB.Recordset
    
    If KeyAscii = 13 Then
        Call cmdStartFind_Click
        Exit Sub
    End If
    
    mstrFindKey = cboFindWay.Text
    
    If mstrFindKey = "�����" Then  '�����
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    ElseIf mstrFindKey = "���￨��" Or mstrFindKey = "����" Then     '���￨��,'������Ҳ��ˢ��
            If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
            
            blnCard = zlCommFun.InputIsCard(txtFindData, KeyAscii, glngSys)
            If blnCard And Len(txtFindData.Text) = gbyt�ſ� - 1 And KeyAscii <> 8 Then
            
                txtFindData.Text = txtFindData.Text & Chr(KeyAscii)
                txtFindData.SelStart = Len(txtFindData.Text)
                
                KeyAscii = 0
                
                Call lvwQueueData.ListItems.Clear
                
                Set rsData = FindQueueData("���￨��", txtFindData.Text)
                
                If rsData Is Nothing Then
                    MsgBox "û�м������������ݡ�", vbInformation, "�Ŷӽк�ϵͳ"
                    Exit Sub
                End If
                
                If rsData.RecordCount <= 0 Then
                    MsgBox "û�м������������ݡ�", vbInformation, "�Ŷӽк�ϵͳ"
                    Exit Sub
                End If
                
                Call LoadDataToFace(lvwQueueData, rsData, "ID")
                
            End If
    ElseIf mstrFindKey = "ҽ����" Then    'ҽ����
    End If
End Sub
