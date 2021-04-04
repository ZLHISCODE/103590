VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmSelect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ѡ����"
   ClientHeight    =   5940
   ClientLeft      =   3840
   ClientTop       =   3525
   ClientWidth     =   5385
   Icon            =   "FrmSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4200
      TabIndex        =   2
      Top             =   630
      Width           =   1100
   End
   Begin VB.CommandButton Cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4200
      TabIndex        =   1
      Top             =   150
      Width           =   1100
   End
   Begin MSComctlLib.TreeView Tvw 
      Height          =   5925
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   10451
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imgTree"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList imgTree 
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
            Picture         =   "FrmSelect.frx":0E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSelect.frx":2B4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSelect.frx":4858
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public BlnSuccess As Boolean '�ɹ���־
Public TreeRec As New ADODB.Recordset
Public CurrentID As Long
Public CurrentName As String
Public StrNode As String
Public lngMode As Integer

Private Sub Cmdȡ��_Click()
    BlnSuccess = False
    Me.Hide
End Sub

Private Sub Cmdȷ��_Click()
    If Tvw.SelectedItem.Tag <> 1 Then Exit Sub
    
    BlnSuccess = True
    Me.Hide
End Sub

Private Sub Form_Load()
    CurrentID = 0
    CurrentName = ""
    LoadInTree
End Sub

Private Function LoadInTree()
    Dim strID As String
    
    Tvw.Nodes.Clear
'    Tvw.Nodes.Add , , "R", StrNode, 1, 1
'    Tvw.Nodes("R").Tag = "R"
    If TreeRec.RecordCount = 0 Then Exit Function
    
    With TreeRec
        If lngMode = 0 Then
            Do While Not .EOF
                If IsNull(!�ϼ�ID) Then
                    If !ĩ�� = 1 Then
                        Tvw.Nodes.Add , , "K_" & !Id, "[" & !���� & "]" & !����, 2, 2
                    Else
                        Tvw.Nodes.Add , , "K_" & !Id, "[" & !���� & "]" & !����, 3, 3
                    End If
                    strID = strID & !Id & ";"
                Else
                    If InStr(strID, !�ϼ�ID & ";") = 0 Then
                        Tvw.Nodes.Add , , "K_" & !Id, "[" & !���� & "]" & !����, 2, 2
                    Else
                        If !ĩ�� = 1 Then
                            Tvw.Nodes.Add "K_" & !�ϼ�ID, 4, "K_" & !Id, "[" & !���� & "]" & !����, 2, 2
                        Else
                            Tvw.Nodes.Add "K_" & !�ϼ�ID, 4, "K_" & !Id, "[" & !���� & "]" & !����, 3, 3
                        End If
                    End If
                End If
                Tvw.Nodes("K_" & !Id).Tag = !ĩ��
                .MoveNext
            Loop
        Else
            Do While Not .EOF
                Tvw.Nodes.Add , , "K_" & !����, "[" & !���� & "]" & !����, 2, 2
                Tvw.Nodes("K_" & !����).Tag = 1
                .MoveNext
            Loop
        End If
    End With
    Tvw.Nodes(1).Selected = True
    tvw_NodeClick Tvw.Nodes(1)
    
'    Tvw.Nodes("R").Selected = True
'    Tvw.Nodes("R").Expanded = True
End Function

Private Sub Tvw_DblClick()
    tvw_NodeClick Tvw.SelectedItem
    Cmdȷ��_Click
End Sub

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
    If Tvw.SelectedItem.Key <> "R" And Tvw.SelectedItem.Children = 0 Then
        CurrentID = Mid(Tvw.SelectedItem.Key, 3)
        CurrentName = Mid(Tvw.SelectedItem, InStr(2, Tvw.SelectedItem, "]") + 1)
    Else
        CurrentID = 0
        CurrentName = ""
    End If
End Sub
