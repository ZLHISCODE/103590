VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServerSelect 
   BorderStyle     =   0  'None
   ClientHeight    =   3390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4335
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imgServer 
      Left            =   3720
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServerSelect.frx":0000
            Key             =   "Server"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwServer 
      Height          =   3015
      Left            =   30
      TabIndex        =   0
      Top             =   90
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgServer"
      SmallIcons      =   "imgServer"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "服务器"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "主机名"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "实例名"
         Object.Width           =   1764
      EndProperty
   End
End
Attribute VB_Name = "frmServerSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrServer As String
Dim mintColumn As Integer

Private Sub Form_Resize()
    With lvwServer
        .Left = ScaleLeft
        .Top = ScaleTop
        .Width = ScaleWidth
        .Height = ScaleHeight
    End With
End Sub

Private Sub lvwServer_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then '仍是刚才那列
        lvwServer.SortOrder = IIf(lvwServer.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwServer.SortKey = mintColumn
        lvwServer.SortOrder = lvwAscending
    End If
End Sub

Public Function GetServer(colServer As Collection, ByVal sngLeft As Single, ByVal sngTop As Single, ByVal strServer As String) As String
'得到用户选择的服务器
    Dim lst As ListItem
    Dim varItem As Variant
    
    mstrServer = ""
    
    lvwServer.ListItems.Clear
    For Each varItem In colServer
        Set lst = lvwServer.ListItems.Add(, , varItem(0), "Server", "Server")
        lst.SubItems(1) = varItem(1)
        lst.SubItems(2) = varItem(2)
        
        If UCase(varItem(0)) = UCase(strServer) Then
            lst.Selected = True
            lst.EnsureVisible
        End If
    Next
    If lvwServer.ListItems.count > 0 And lvwServer.SelectedItem Is Nothing Then
       lvwServer.ListItems(1).Selected = True
    End If
    
    Left = sngLeft
    Top = sngTop
    Me.Show vbModal, frmUserLogin
    '返回值
    GetServer = mstrServer
End Function

Private Sub lvwServer_DblClick()
    If Not lvwServer.SelectedItem Is Nothing Then
        mstrServer = lvwServer.SelectedItem.Text
    End If
    Unload Me
End Sub

Private Sub lvwServer_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Not lvwServer.SelectedItem Is Nothing Then
            mstrServer = lvwServer.SelectedItem.Text
        End If
        Unload Me
    ElseIf KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub
