VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����Աά��"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton Command5 
      Caption         =   "�˳�(&E)"
      Height          =   405
      Left            =   4470
      TabIndex        =   5
      Top             =   1980
      Width           =   1100
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�ϴ�(&T)"
      Height          =   405
      Left            =   4470
      TabIndex        =   4
      Top             =   1524
      Width           =   1100
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�޸�(&M)"
      Height          =   405
      Left            =   4470
      TabIndex        =   3
      Top             =   618
      Width           =   1100
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ɾ��(&D)"
      Height          =   405
      Left            =   4470
      TabIndex        =   2
      Top             =   1071
      Width           =   1100
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���(&A)"
      Height          =   405
      Left            =   4470
      TabIndex        =   1
      Top             =   165
      Width           =   1100
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4485
      Left            =   152
      TabIndex        =   0
      Top             =   135
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   7911
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "���"
         Object.Width           =   1005
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "����"
         Object.Width           =   1640
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "����"
         Object.Width           =   1640
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Ȩ��"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmUserList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    If frmUserEdit.userEdit(0) = True Then fillList
End Sub

Private Sub Command2_Click()
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    If ListView1.SelectedItem.Text = "0" Then
        MsgBox "����ɾ���̶�����Ա��Ȩ��", vbInformation, "ɾ��"
        Exit Sub
    End If
    If MsgBox("��ȷ���Ƿ�ɾ���˲���Ա��ҽ��Ȩ�ޣ�", vbQuestion + vbYesNo, "ɾ��") = vbNo Then Exit Sub
    gcn��ͨ.Execute "Delete From tab_czry where oper=" & ListView1.SelectedItem.Text
    ListView1.ListItems.Remove ListView1.SelectedItem.Index
End Sub

Private Sub Command3_Click()
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    If frmUserEdit.userEdit(1, ListView1.SelectedItem.Text) = True Then fillList
End Sub

Private Sub Command4_Click()
    Dim rsTemp As New ADODB.Recordset, strPara As String
    Set rsTemp = gcn��ͨ.Execute("Select * From tab_czry order by oper")
    While Not rsTemp.EOF
        If rsTemp(0) <> 0 Then strPara = strPara & ";" & rsTemp!oper & "," & rsTemp!Name & "," & rsTemp!password & "," & rsTemp!popedom
        rsTemp.MoveNext
    Wend
    strPara = Mid(strPara, 2)
    If strPara = "" Then strPara = " "
    
    frmConn��ͨ.Execute "I050", 5, strPara, "�����ϴ�����Ա����......"
    
End Sub

Private Sub Command5_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    fillList
End Sub

Public Sub fillList()
    Dim rsTemp As New ADODB.Recordset, lstTemp As ListItem
    ListView1.ListItems.Clear
    Set rsTemp = gcn��ͨ.Execute("Select * from tab_czry order by oper")
    While Not rsTemp.EOF
        Set lstTemp = ListView1.ListItems.Add(, "K" & rsTemp!hisid, rsTemp!oper)
        lstTemp.ListSubItems.Add , , rsTemp!Name
        lstTemp.ListSubItems.Add , , rsTemp!password
        lstTemp.ListSubItems.Add , , IIf(rsTemp!popedom = 2, "ϵͳ����", IIf(rsTemp!popedom = 10000, "����ҵ��", "סԺҵ��"))
    
        rsTemp.MoveNext
    Wend

End Sub
