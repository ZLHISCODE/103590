VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTreeLeafSel 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4890
   Icon            =   "frmTreeLeafSel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3600
      TabIndex        =   2
      Top             =   600
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   3600
      TabIndex        =   1
      Top             =   150
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
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   3630
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTreeLeafSel.frx":0442
            Key             =   "Write"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTreeLeafSel.frx":0896
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTreeLeafSel.frx":0CEA
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTreeLeafSel.frx":113E
            Key             =   "Dept"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTreeLeafSel.frx":1458
            Key             =   "Man"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTreeLeafSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrID As String
Dim mstr���� As String
Dim mblnSecceed As Boolean
Dim mblnNode As Boolean
Dim mstrCaption As String

Private Sub cmdCancel_Click()
    mblnSecceed = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If tvw.SelectedItem.Image <> "Item" And tvw.SelectedItem.Image <> "Man" Then Exit Sub
    mblnSecceed = True
    With tvw.SelectedItem
        mstrID = Mid(.Key, 2)
        mstr���� = .Text
    End With
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = mstrCaption
    RestoreWinState Me
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

Public Function ShowTree(ByVal strSQL As String, strID As String, str���� As String, _
                ByVal strCaption As String, Optional ByVal bln��Ա As Boolean = False) As Boolean
'����:����SQL�����ʾ������Ŀ,��ѡ��ĳ��ĩ����Ŀ
'����:strSql    SQL���
'     strID     ������ѡ����Ŀ��ID
'     str����   ������ѡ����Ŀ������
'     strCaption   ���ڵı���
'����:����ѡ�񷵻�True,���򷵻�False.
    Dim rsTree As New ADODB.Recordset
    Dim strTemp As String, strPre As String
    Dim nod As Node
    
    On Error GoTo errHandle
    mstrID = strID
    mstrCaption = strCaption
'    Call SQLTest(App.ProductName, Me.Caption, strSQL)
    Set rsTree = zlDatabase.OpenSQLRecord(strSQL, "ShowTree")
'    Call SQLTest
    If rsTree.RecordCount = 0 Then
        MsgBox "ѡ����û�ҵ�������Ŀ��", vbExclamation, gstrSysName
        ShowTree = False
        Exit Function
    End If
    tvw.Nodes.Clear
    Do Until rsTree.EOF
        If bln��Ա = True Then
            strTemp = IIF(rsTree("ĩ��") = 1, "Man", "Dept")
        Else
            strTemp = IIF(rsTree("ĩ��") = 1, "Item", "Write")
        End If
        
        If IsNull(rsTree("�ϼ�id")) Then
            Set nod = tvw.Nodes.Add(, , "C" & rsTree("id"), rsTree("����"), strTemp, strTemp)
        Else
            strPre = IIF(rsTree("ĩ��") = 1, "K", "C")
            tvw.Nodes.Add "C" & rsTree("�ϼ�id"), tvwChild, strPre & rsTree("id"), rsTree("����"), strTemp, strTemp
        End If
        nod.Sorted = True
        rsTree.MoveNext
    Loop
    
    If strID <> "" And strID <> "0" Then
        '���ܸýڵ��Ѿ���ɾ����
        On Error Resume Next
        tvw.Nodes("K" & strID).Selected = True
        tvw.Nodes("K" & strID).EnsureVisible
    End If
    Me.Show vbModal
    ShowTree = mblnSecceed
    '�ɹ��˲ŷ���ֵ
    If mblnSecceed = True Then
        strID = mstrID
        str���� = mstr����
    End If
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me
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
