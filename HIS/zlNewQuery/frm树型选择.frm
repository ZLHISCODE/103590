VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm����ѡ�� 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4890
   Icon            =   "frm����ѡ��.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   3690
      TabIndex        =   1
      Top             =   210
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3690
      TabIndex        =   0
      Top             =   660
      Width           =   1100
   End
   Begin MSComctlLib.TreeView tvw 
      Height          =   3315
      Left            =   120
      TabIndex        =   2
      Top             =   450
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   5847
      _Version        =   393217
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
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm����ѡ��.frx":0442
            Key             =   "Write"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm����ѡ��.frx":0896
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm����ѡ��.frx":0CEA
            Key             =   "End"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm����ѡ��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Dim mstrID As String
Dim mstr�ϼ�ID As String
Dim mstr�ϼ����� As String
Dim mstr�ϼ����� As String
Dim mstrԭ���� As String
Dim mblnSecceed As Boolean
Dim mblnNode As Boolean
Dim mblnRoot As Boolean '����ѡ������
Dim mblnSelĩ�� As Boolean

Dim mstrCaption As String

Private Sub cmdCancel_Click()
    mblnSecceed = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim nod As Node
    Dim i As Integer
    Dim str���� As String
    
    
    Set nod = tvw.SelectedItem
    
    If mblnSelĩ�� And nod.Tag <> "1" Then Exit Sub
    
    If mstrԭ���� <> "" Then
        If nod.Key = "Root" Then
            str���� = ""
        Else
            str���� = Mid(nod.Text, 2, InStr(nod.Text, "��") - 2)
        End If
        'mstrIDΪ�ձ�ʾ��������ʱ���������¼���ϵ
        If mstrԭ���� = Mid(str����, 1, Len(mstrԭ����)) And mstrID <> "" Then Exit Sub
    End If
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
            If mblnRoot = False Then Exit Sub
            mstr�ϼ�ID = ""
            mstr�ϼ����� = "��"
            mstr�ϼ����� = ""
        Else
            i = InStr(.Text, "��")
            mstr�ϼ�ID = Mid(.Key, 2)
            mstr�ϼ����� = Mid(.Text, i + 1)
            mstr�ϼ����� = Mid(.Text, 2, i - 2)
        End If
    End With
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = mstrCaption
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

Public Function ShowTree(ByVal strSQL As String, str�ϼ�ID As String, str�ϼ����� As String, str�ϼ����� As String, ByVal strID As String, ByVal strCaption As String, _
    ByVal strRoot As String, Optional blnRoot As Boolean = True, Optional strԭ���� As String, Optional blnSelĩ�� As Boolean = False) As Boolean
'����:����SQL�����ʾ������Ŀ,��ѡ��ĳ��ĩ����Ŀ
'����:strSql        SQL���
'     str�ϼ�ID     ������ѡ����Ŀ���ϼ�ID
'     str�ϼ�����   ������ѡ����Ŀ���ϼ�����
'     str�ϼ�����   ������ѡ����Ŀ���ϼ�����
'     strID         ������ѡ����Ŀ��ID
'     strRoot       �����ı���
'     strICO        ͼ����Դ������
'     strCaption    ���ڵı���
'����:����ѡ�񷵻�True,���򷵻�False.
    
    Dim rs���� As New ADODB.Recordset
    
    mblnRoot = blnRoot
    mstrCaption = strCaption
    
    Set rs���� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    mblnSelĩ�� = blnSelĩ��
    
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
        If blnSelĩ�� Then
            tvw.Nodes("C" & rs����("id")).Tag = Val(IIf(IsNull(rs����("ĩ��")), "0", rs����("ĩ��")))
            If tvw.Nodes("C" & rs����("id")).Tag = "1" Then
                tvw.Nodes("C" & rs����("id")).Image = "End"
                tvw.Nodes("C" & rs����("id")).SelectedImage = "End"
            End If
        End If
        rs����.MoveNext
    Loop
    If str�ϼ�ID = "" Then
        tvw.Nodes("Root").Selected = True
        tvw.Nodes("Root").Expanded = True
    Else
        tvw.Nodes("C" & str�ϼ�ID).Selected = True
        tvw.Nodes("C" & str�ϼ�ID).EnsureVisible
    End If
    
    mstrID = strID
    mstrԭ���� = strԭ����
    Me.Show vbModal
    ShowTree = mblnSecceed
    '�ɹ��˲ŷ���ֵ
    If mblnSecceed = True Then
        str�ϼ�ID = mstr�ϼ�ID
        str�ϼ����� = mstr�ϼ�����
        str�ϼ����� = mstr�ϼ�����
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub tvw_DblClick()
    If mblnNode Then cmdOK_Click
End Sub

Private Sub tvw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnNode = False
End Sub

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
    mblnNode = True
End Sub
