VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm��Ӧ��ѡ�� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��Ӧ��ѡ��"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5325
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   4155
      TabIndex        =   2
      Top             =   60
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4155
      TabIndex        =   1
      Top             =   435
      Width           =   1100
   End
   Begin MSComctlLib.TreeView tvwList 
      Height          =   4410
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   7779
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   4380
      Top             =   1695
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
            Picture         =   "frm��Ӧ��ѡ��.frx":0000
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ��ѡ��.frx":0458
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ��ѡ��.frx":08B0
            Key             =   "ItemNo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ��ѡ��.frx":0D04
            Key             =   "No"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ��ѡ��.frx":115C
            Key             =   "Write"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm��Ӧ��ѡ��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mselStr As String
Private mstrPrivs As String
Private msngDownX As Single, msngDownY As Single

Public Function SelDept(ByVal strPrivs As String) As String
    '-----------------------------------------------------------------------------------------------------------
    '����:��Ӧ��ѡ��
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-08-18 14:46:41
    '-----------------------------------------------------------------------------------------------------------
    Dim rstTemp As New ADODB.Recordset, strSql As String
    mstrPrivs = strPrivs
    
    tvwList.Nodes.Clear
    tvwList.Nodes.Add , , "Root", "���й�Ӧ��", 1
    Set tvwList.SelectedItem = tvwList.Nodes("Root")
    tvwList.SelectedItem.Expanded = True
    tvwList.SelectedItem.Sorted = True
    Dim strȨ�� As String
    
    strȨ�� = " and (ĩ��<>1 or (ĩ��=1 " & zl_��ȡվ������() & "  and " & Get����Ȩ��(mstrPrivs) & ")) "
    strSql = "" & _
        "   Select ID,�ϼ�ID,����,����,ĩ�� " & _
        "   From ��Ӧ�� " & _
        "   Where (����ʱ�� is null or  ����ʱ��=TO_DATE('3000-1-1','yyyy-MM-dd'))" & strȨ�� & _
        "   start with �ϼ�ID is null connect by prior ID =�ϼ�ID"
    On Error GoTo errHandle
    zlDatabase.OpenRecordset rstTemp, strSql, Me.Caption
    
    While Not rstTemp.EOF
        If IsNull(rstTemp!�ϼ�ID) Then
            tvwList.Nodes.Add "Root", tvwChild, "P" & rstTemp!ID, "[" & rstTemp!���� & "]" & rstTemp!����, IIf(rstTemp!ĩ�� <> 1, 5, 2)
        Else
            tvwList.Nodes.Add "P" & rstTemp!�ϼ�ID, tvwChild, "P" & rstTemp!ID, "[" & rstTemp!���� & "]" & rstTemp!����, IIf(rstTemp!ĩ�� <> 1, 5, 2)
        End If
        rstTemp.MoveNext
    Wend
    Me.Show vbModal
    SelDept = mselStr
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then Resume
End Function

Private Sub cmdCancel_Click()
    mselStr = ""
    Me.Hide
End Sub


Private Sub cmdOK_Click()
    If tvwList.SelectedItem.Image = 2 Then
        mselStr = Mid(tvwList.SelectedItem.Key, 2) & "," & tvwList.SelectedItem.Text
        Me.Hide
    End If
End Sub

Private Sub tvwList_DblClick()
    If tvwList.HitTest(msngDownX, msngDownY) Is Nothing Then Exit Sub
    If tvwList.SelectedItem.Image = 2 Then
        mselStr = Mid(tvwList.SelectedItem.Key, 2) & "," & Mid(tvwList.SelectedItem.Text, InStr(tvwList.SelectedItem.Text, "��") + 1)
        Me.Hide
    End If
End Sub

Private Sub tvwList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    msngDownX = X
    msngDownY = Y
End Sub

Private Sub tvwList_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Image = 2 Then
        cmdOK.Enabled = True
    Else
        cmdOK.Enabled = False
    End If
End Sub
