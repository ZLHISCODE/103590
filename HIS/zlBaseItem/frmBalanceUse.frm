VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBalanceUse 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���㷽ʽӦ������"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   Icon            =   "frmBalanceUse.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   3780
      TabIndex        =   1
      Top             =   210
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3780
      TabIndex        =   2
      Top             =   660
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   3780
      TabIndex        =   3
      Top             =   3360
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvw���㷽ʽ 
      Height          =   2925
      Left            =   90
      TabIndex        =   0
      Top             =   1230
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   5159
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "_Ӧ�ó���"
         Object.Tag             =   "Ӧ�ó���"
         Text            =   "Ӧ�ó���"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "_ȱʡ��־"
         Object.Tag             =   "ȱʡ��־"
         Text            =   "ȱʡ��־"
         Object.Width           =   1058
      EndProperty
   End
   Begin VB.Label lbl��ʾ 
      Caption         =   $"frmBalanceUse.frx":000C
      Height          =   915
      Left            =   120
      TabIndex        =   4
      Top             =   180
      Width           =   3285
   End
End
Attribute VB_Name = "frmBalanceUse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstr���� As String
Dim mblnItem As Boolean
Dim mintSuccess As Integer
Dim mblnChange As Boolean     '�Ƿ�ı���

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, 5
End Sub

Private Sub cmdOK_Click()
    If Save����() = False Then Exit Sub
    mintSuccess = mintSuccess + 1
    mblnChange = False
    Unload Me
End Sub

Private Function Save����() As Boolean
'����:����༭�����ݽ��㷽ʽ����
'����:
'����ֵ:�ɹ�����True,����ΪFalse
    Dim i As Integer
    Dim str���� As String
    On Error GoTo ErrHandle
    '������ѡ�еĹ�����������һ����
    '������ѡ�еĹ�����������һ����
    For i = 1 To lvw���㷽ʽ.ListItems.Count
        If lvw���㷽ʽ.ListItems(i).Checked = True Then
            str���� = str���� & lvw���㷽ʽ.ListItems(i) & ":"
            str���� = str���� & IIF(lvw���㷽ʽ.ListItems(i).SubItems(1) = "", "0;", "1;")
        End If
    Next
    
    '�޸�
    gstrSQL = "zl_���㷽ʽӦ��_update( '" & mstr���� & "','" & str���� & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    mblnChange = False
    Save���� = True
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function �༭����(ByVal str���� As String) As Boolean
'����:��������õĽ��㷽ʽ�����ڽ���ͨѶ�ĳ���
'����:str����     ��ǰ�༭�Ľ��㷽ʽ�ı���
'����ֵ:�༭�ɹ�����True,����ΪFalse
    Dim rs���㷽ʽ As New ADODB.Recordset
    Dim ObjItem As Object
    mblnChange = False
    mintSuccess = 0
        
    mstr���� = str����
    lbl��ʾ.Caption = Replace(lbl��ʾ.Caption, "�����㷽ʽ", "�ڡ�" & str���� & "�����㳡����")
    
    '�������㳡��:���տ�ֻӦ����Ԥ����
    '�����շ�,��������,����Ӧ����Ӧ����:33722
    '75134:���ϴ�,2014/7/14,�ų�����Ϊ9�Ľ��㷽ʽ
    '82990:���ϴ�,2015/3/9,ҽ���������ڲ�����
    On Error GoTo ErrHandle
    gstrSQL = "Select A.����,A.����,B.���㷽ʽ,B.ȱʡ��־,nvl(A.Ӧ����,0) as Ӧ����" & _
        " From ���㷽ʽ A,���㷽ʽӦ�� B" & _
        " Where A.����=B.���㷽ʽ(+) And B.Ӧ�ó���(+)=[1] " & _
        IIF(str���� <> "Ԥ����", " And A.����<>5", "") & _
        IIF(str���� = "������" Or str���� = "���ѿ�", " And A.���� In(1,2,8)", "") & _
        IIF(InStr("�շ�,����", str����) = 0, " And nvl(A.Ӧ����,0)<>1 ", "") & _
        " And A.����<>9 And b.���ʽ(+) Is Null" & _
        " Order by A.����"
        
    Set rs���㷽ʽ = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str����)
        
    lvw���㷽ʽ.ListItems.Clear
    Do Until rs���㷽ʽ.EOF
       Set ObjItem = lvw���㷽ʽ.ListItems.Add(, "C" & rs���㷽ʽ!����, rs���㷽ʽ!����)
        ObjItem.Tag = Nvl(rs���㷽ʽ!����, 1) & "," & Val(Nvl(rs���㷽ʽ!Ӧ����))
        
        If Not IsNull(rs���㷽ʽ!���㷽ʽ) Then
            ObjItem.Checked = True
            If Nvl(rs���㷽ʽ!����, 1) < 3 Then
                ObjItem.SubItems(1) = IIF(Nvl(rs���㷽ʽ!ȱʡ��־, 0) = 1, "ȱʡ", "")
            End If
        End If
        rs���㷽ʽ.MoveNext
    Loop
    
    frmBalanceUse.Show vbModal
    �༭���� = mintSuccess > 0
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub lvw���㷽ʽ_DblClick()
    If mblnItem = False Then Exit Sub
    Call ChangeServer
    mblnItem = False
End Sub

Private Sub lvw���㷽ʽ_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("C") Or KeyAscii = Asc("c") Then Call ChangeServer
End Sub

Private Sub ChangeServer()
    Dim i As Integer, j As Integer
    Dim varData As Variant
    If lvw���㷽ʽ.SelectedItem Is Nothing Then Exit Sub
    
    With lvw���㷽ʽ.SelectedItem
        If .Checked = False Then Exit Sub
        varData = Split(.Tag & ",", ",")
        If InStr("1,2,7,8", Val(varData(0))) = 0 Then .SubItems(1) = "": Exit Sub 'ҽ�����㼰���տ��Ϊȱʡ���㷽ʽ
        If Val(varData(1)) = 1 Then .SubItems(1) = "": Exit Sub 'Ӧ��������ó�ȱʡ��ʽ
        
        If .SubItems(1) = "" Then
            .SubItems(1) = "ȱʡ"
            mblnChange = True
        Else
            .SubItems(1) = ""
            mblnChange = True
        End If
        cmdOK.Enabled = True
    End With
    
    '��֤Ψһ��
    If lvw���㷽ʽ.SelectedItem.SubItems(1) <> "" Then
        j = lvw���㷽ʽ.SelectedItem.Index
        For i = 1 To lvw���㷽ʽ.ListItems.Count
            If i <> j Then
                lvw���㷽ʽ.ListItems(i).SubItems(1) = ""
            End If
        Next
    End If
End Sub

Private Sub lvw���㷽ʽ_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    cmdOK.Enabled = True
    mblnChange = True
    If Item.Checked = False And Item.SubItems(1) = "ȱʡ" Then Item.SubItems(1) = ""
End Sub

Private Sub lvw���㷽ʽ_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mblnItem = True
End Sub
