VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm�շ�ϸĿѡ�� 
   Caption         =   "�շ�ϸĿ"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5130
   Icon            =   "frm�շ�ϸĿѡ��.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   5130
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3690
      TabIndex        =   2
      Top             =   210
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3690
      TabIndex        =   1
      Top             =   750
      Width           =   1100
   End
   Begin MSComctlLib.TreeView tvw 
      Height          =   3315
      Left            =   120
      TabIndex        =   0
      Top             =   450
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   5847
      _Version        =   393217
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�շ�ϸĿѡ��.frx":0442
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�շ�ϸĿѡ��.frx":0894
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�շ�ϸĿѡ��.frx":0CE6
            Key             =   "Write"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm�շ�ϸĿѡ��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrID As String
Dim mstr���� As String
Dim mstr���� As String
Dim mblnSecceed As Boolean
Dim mblnNode As Boolean

Private Sub cmdCancel_Click()
    mblnSecceed = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If tvw.SelectedItem.Image <> "Item" Then Exit Sub
    mblnSecceed = True
    With tvw.SelectedItem
        mstrID = Mid(.Key, 2)
        mstr���� = Mid(.Text, 2, InStr(.Text, "��") - 2)
        mstr���� = Mid(.Text, InStr(.Text, "��") + 1)
    End With
    Unload Me
End Sub

Private Sub Form_Load()
    RestoreWinState Me, App.ProductName
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    tvw.Top = 100
    tvw.Left = 100
    
    tvw.Height = ScaleHeight - 200
    If Me.ScaleWidth > 3000 Then
        cmdOK.Left = ScaleWidth - cmdOK.Width - 200
        cmdCancel.Left = cmdOK.Left
'        cmdHelp.Left = cmdOK.Left
        tvw.Width = cmdOK.Left - tvw.Left - 200
    End If
End Sub

Public Function ShowTree(strID As String, str���� As String, str���� As String) As Boolean
'����:��ʾ�����շ�ϸĿ,���ó�ѡ��
'����:strID     ������ѡ���շ�ϸĿ��ID
'     str����   ������ѡ���շ�ϸĿ������
'����:����ѡ�񷵻�True,���򷵻�False.

    Dim rsTree As New ADODB.Recordset
    Dim strTemp As String
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    mstrID = strID
    
    strSQL = "select ����,��� from �շ����"
    Call OpenRecordset(rsTree, Me.Caption, strSQL)
    
    tvw.Nodes.Clear
    Do Until rsTree.EOF
        tvw.Nodes.Add , , "C" & rsTree("����"), "��" & rsTree("����") & "��" & rsTree("���"), "Root", "Root"
        tvw.Nodes.Add "C" & rsTree("����"), tvwChild, "K" & rsTree("����"), "��ʱ"
        rsTree.MoveNext
    Loop
    Me.Show vbModal
    ShowTree = mblnSecceed
    '�ɹ��˲ŷ���ֵ
    If mblnSecceed = True Then
        strID = mstrID
        str���� = mstr����
        str���� = mstr����
    End If
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub tvw_DblClick()
    If Not mblnNode Then Exit Sub
    cmdOK_Click
End Sub

Private Sub tvw_Expand(ByVal Node As MSComctlLib.Node)
    Dim rsTree As New ADODB.Recordset
    Dim strTemp As String
    Dim strSQL As String
    
    If Node.Image = "Root" And Left(Node.Child.Key, 1) = "K" Then
    'ֻ��δ�����¼��ĵĸ��ڵ㴦��
        
        'ɾ����ʱ�ڵ�
        tvw.Nodes.Remove Node.Child.Key
        
        '�������µ��¼�
        rsTree.CursorLocation = adUseClient
        strSQL = "select ID,�ϼ�ID,����,����,ĩ�� from �շ�ϸĿ  " & _
            " where ����ʱ�� is null or ����ʱ�� =to_date('3000-01-01','YYYY-MM-DD') " & _
            " start with �ϼ�ID is null and ���='" & Mid(Node.Key, 2, 1) & "' connect by prior ID =�ϼ�ID"
        Call OpenRecordset(rsTree, Me.Caption, strSQL)
        
        Do Until rsTree.EOF
            strTemp = IIf(rsTree("ĩ��") = 1, "Item", "Write")
            If IsNull(rsTree("�ϼ�id")) Then
                tvw.Nodes.Add Node.Key, tvwChild, "_" & rsTree("id"), "��" & rsTree("����") & "��" & rsTree("����"), strTemp, strTemp
            Else
                tvw.Nodes.Add "_" & rsTree("�ϼ�id"), tvwChild, "_" & rsTree("id"), "��" & rsTree("����") & "��" & rsTree("����"), strTemp, strTemp
            End If
            rsTree.MoveNext
        Loop
    End If
End Sub

Private Sub tvw_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call cmdOK_Click
End Sub

Private Sub tvw_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnNode = False
End Sub

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
    mblnNode = True
End Sub
