VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm����ѡ�� 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "����ѡ��"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4890
   ControlBox      =   0   'False
   Icon            =   "frm����ѡ��.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3690
      TabIndex        =   2
      Top             =   30
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3690
      TabIndex        =   1
      Top             =   450
      Width           =   1100
   End
   Begin MSComctlLib.TreeView tvw 
      Height          =   4305
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   7594
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      Sorted          =   -1  'True
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
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm����ѡ��.frx":0442
            Key             =   "Write"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm����ѡ��.frx":0896
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm����ѡ��.frx":0CEA
            Key             =   "Read"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm����ѡ��.frx":113E
            Key             =   "Root"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm����ѡ��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrID As String
Dim mstr�ϼ�ID As String
Dim mstr�ϼ����� As String
Dim mstr�ϼ����� As String
Dim mblnSecceed As Boolean
Dim mblnNode As Boolean

Private Sub cmdCancel_Click()
    mblnSecceed = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim nod As Node
    Dim intTemp As Integer
    
    Set nod = tvw.SelectedItem
    Do Until nod.Key = "Root"
        If mstrID = Mid(nod.Key, 2) Then
'            MsgBox "�˽ڵ㲻����Ҫ������ѡ��", vbExclamation, gstrSysName
            Exit Sub
        End If
        Set nod = nod.Parent
    Loop
    mblnSecceed = True
    With tvw.SelectedItem
        If .Key = "Root" Then
            mstr�ϼ�ID = ""
            mstr�ϼ����� = "��"
            mstr�ϼ����� = ""
        Else
            intTemp = InStr(.Text, "��")
            mstr�ϼ�ID = Mid(.Key, 2)
            mstr�ϼ����� = Mid(.Text, intTemp + 1)
            mstr�ϼ����� = Mid(.Text, 2, intTemp - 2)
        End If
    End With
    Unload Me
End Sub

Public Function ShowTree(ByVal strSql As String, str�ϼ�ID As String, str�ϼ����� As String, str�ϼ����� As String, strID As String, ByVal strCaption As String, ByVal strRoot As String) As Boolean
    Dim rs���� As New ADODB.Recordset
    
    mstrID = strID
    
    On Error GoTo errHandle
    zlDatabase.OpenRecordset rs����, strSql, Me.Caption
    
    tvw.Nodes.Clear
    tvw.Nodes.Add , , "Root", strRoot, "Root", "Root"
    tvw.Nodes("Root").Sorted = True
    Do Until rs����.EOF
        
        If IsNull(rs����("�ϼ�id")) Then
            tvw.Nodes.Add "Root", tvwChild, "C" & rs����("id"), "��" & rs����("����") & "��" & rs����("����"), "Write", "Write"
        Else
            tvw.Nodes.Add "C" & rs����("�ϼ�id"), tvwChild, "C" & rs����("id"), "��" & rs����("����") & "��" & rs����("����"), "Write", "Write"
        End If
        tvw.Nodes("C" & rs����("id")).Sorted = True
        rs����.MoveNext
    Loop
    If str�ϼ�ID = "" Then
        tvw.Nodes("Root").Selected = True
        tvw.Nodes("Root").Expanded = True
    Else
        tvw.Nodes("C" & str�ϼ�ID).Selected = True
        tvw.Nodes("C" & str�ϼ�ID).EnsureVisible
    End If
    Me.Caption = strCaption
    Me.Show vbModal
    ShowTree = mblnSecceed
    '�ɹ��˲ŷ���ֵ
    If mblnSecceed = True Then
        str�ϼ�ID = mstr�ϼ�ID
        str�ϼ����� = mstr�ϼ�����
        str�ϼ����� = mstr�ϼ�����
    End If
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then Resume
End Function

Private Sub tvw_DblClick()
    If mblnNode Then cmdOK_Click
End Sub

Private Sub tvw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnNode = False
End Sub

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
    mblnNode = True
End Sub
