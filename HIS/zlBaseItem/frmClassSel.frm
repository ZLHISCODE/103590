VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClassSel 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "�����������"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4890
   Icon            =   "frmClassSel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
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
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   3570
      Top             =   2700
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClassSel.frx":0442
            Key             =   "Write"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClassSel.frx":0896
            Key             =   "Root"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmClassSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrID As String
Dim mstr�ϼ�ID As String
Dim mstr�ϼ����� As String
Dim mstr���뷶Χ As String
Dim mblnRoot As Boolean '����ѡ������

Dim mblnNode As Boolean
Dim mblnSecceed As Boolean

Private Sub cmdCancel_Click()
    mblnSecceed = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim nod As Node
    Dim i As Integer
    Dim str���� As String
    
    Set nod = tvw.SelectedItem
    
    '�ж��Ƿ񱾼�
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
            mstr���뷶Χ = ""
        Else
            mstr�ϼ�ID = Mid(.Key, 2)
            mstr�ϼ����� = .Text
            mstr���뷶Χ = .Tag
        End If
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
        tvw.Width = cmdOK.Left - tvw.Left - 200
    End If
End Sub

Public Function ShowTree(str�ϼ�ID As String, str�ϼ����� As String, str���뷶Χ As String _
    , ByVal str������� As String, ByVal strID As String, Optional blnRoot As Boolean = True) As Boolean
'����:����SQL�����ʾ������Ŀ,��ѡ��ĳ��ĩ����Ŀ
'����:str�ϼ�ID     ������ѡ����Ŀ���ϼ�ID
'     str�ϼ�����   ������ѡ����Ŀ���ϼ�����
'     strID         ��ѡ����Ŀ��ID�������ж��Ƿ������¼���
'     blnRoot       �Ƿ�����ѡ����ڵ�
'����:����ѡ�񷵻�True,���򷵻�False.
    
    Dim rsTemp As New ADODB.Recordset
    Dim nodTemp As Node
    
    mblnRoot = blnRoot
    mstrID = strID
    
    On Error GoTo ErrHandle
    gstrSQL = "select level,ID,�ϼ�ID,���,����,���뷶Χ from ����������� where ���=[1] " & _
        " start with �ϼ�ID is null connect by prior id=�ϼ�ID order by level,���"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, str�������)
        
    tvw.Nodes.Clear
    tvw.Nodes.Add , , "Root", "�������", "Root", "Root"
    Do Until rsTemp.EOF
        If IsNull(rsTemp("�ϼ�ID")) Then
            Set nodTemp = tvw.Nodes.Add("Root", tvwChild, "K" & rsTemp("ID"), "��" & rsTemp("���") & "��" & Trim(rsTemp("����")), "Write", "Write")
        Else
            Set nodTemp = tvw.Nodes.Add("K" & rsTemp("�ϼ�ID"), tvwChild, "K" & rsTemp("ID"), "��" & rsTemp("���") & "��" & Trim(rsTemp("����")), "Write", "Write")
        End If
        nodTemp.Tag = IIF(IsNull(rsTemp("���뷶Χ")), "", rsTemp("���뷶Χ"))
        rsTemp.MoveNext
        
    Loop
    If str�ϼ�ID = "" Then
        tvw.Nodes("Root").Selected = True
        tvw.Nodes("Root").Expanded = True
    Else
        tvw.Nodes("K" & str�ϼ�ID).Selected = True
        tvw.Nodes("K" & str�ϼ�ID).EnsureVisible
    End If
    
    Me.Show vbModal
    ShowTree = mblnSecceed
    '�ɹ��˲ŷ���ֵ
    If mblnSecceed = True Then
        str�ϼ�ID = mstr�ϼ�ID
        str�ϼ����� = mstr�ϼ�����
        str���뷶Χ = mstr���뷶Χ
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub tvw_DblClick()
    If mblnNode Then cmdOK_Click
End Sub

Private Sub tvw_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnNode = False
End Sub

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
    mblnNode = True
End Sub
