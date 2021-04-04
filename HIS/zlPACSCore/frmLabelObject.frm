VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLabelObject 
   Caption         =   "当前图象对象分析"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13935
   Icon            =   "frmLabelObject.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   13935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   1085
      ButtonWidth     =   2725
      ButtonHeight    =   926
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "是否显示系统对象"
            Key             =   "DispSystem"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   11520
      Top             =   4788
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6036
      Left            =   840
      TabIndex        =   0
      Top             =   960
      Width           =   10968
      _ExtentX        =   19341
      _ExtentY        =   10636
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "序号"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "类型"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "注释"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "可见"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "左"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "顶"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "宽度"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "高度"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "连接对象"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "文字"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "前景色"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "背景色"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmLabelObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public im As DicomImage
Public f As frmViewer
Private Sub Form_Load()
    load
End Sub

Private Sub Form_Resize()
    ListView1.top = Me.ScaleTop + Me.Toolbar1.height
    ListView1.left = Me.ScaleLeft
    ListView1.height = Me.ScaleHeight - Me.Toolbar1.height
    ListView1.width = Me.ScaleWidth
End Sub

Private Sub ListView1_DblClick()
    Set im = f.SelectedImage
    load
End Sub
Sub load()
    If im Is Nothing Then
        MsgBox "没有选择图像", vbExclamation, gstrSysName
        Exit Sub
    End If
    Dim l As DicomLabel
    Dim o As ListItem
    ListView1.ListItems.Clear
    If f.SelectedLabel Is Nothing Then
        Set o = ListView1.ListItems.Add(, , "SelectedLabel Is Nothing ")
    Else
        Set l = f.SelectedLabel
        Set o = ListView1.ListItems.Add(, , "SelectedLabel IS " & im.Labels.IndexOf(l))
        o.SubItems(1) = l.LabelType & ":" & funLabelType(l)
        o.SubItems(2) = l.Tag
            o.SubItems(3) = IIf(l.Visible, "O", "")
            o.SubItems(4) = l.left
            o.SubItems(5) = l.top
            o.SubItems(6) = l.width
            o.SubItems(7) = l.height
            o.SubItems(9) = l.Text
            o.SubItems(10) = l.ForeColour
            o.SubItems(11) = l.BackColour
            On Error Resume Next
            o.SubItems(8) = im.Labels.IndexOf(l.TagObject)
            On Error GoTo 0
    End If
    For Each l In im.Labels
        If im.Labels.IndexOf(l) > IIf(Me.Toolbar1.Buttons("DispSystem").Value = 1, 0, G_INT_SYS_LABEL_COUNT) Then
            Set o = ListView1.ListItems.Add(, , im.Labels.IndexOf(l))
            o.SubItems(1) = l.LabelType & ":" & funLabelType(l)
            o.SubItems(2) = l.Tag
            o.SubItems(3) = IIf(l.Visible, "O", "")
            o.SubItems(4) = l.left
            o.SubItems(5) = l.top
            o.SubItems(6) = l.width
            o.SubItems(7) = l.height
            o.SubItems(9) = l.Text
            o.SubItems(10) = l.ForeColour
            o.SubItems(11) = l.BackColour
            On Error Resume Next
            o.SubItems(8) = im.Labels.IndexOf(l.TagObject)
            On Error GoTo 0
        End If
    Next
End Sub

Private Sub Timer1_Timer()
    load
    Me.Timer1.Enabled = True
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Key = "DispSystem" Then Button.Value = IIf(Button.Value = 1, 0, 1)
End Sub
