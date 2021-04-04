VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm工伤认定编号选择 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "请选择一种工伤认定编号"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
   ControlBox      =   0   'False
   Icon            =   "frm工伤认定编号选择.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3180
      Top             =   1860
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
            Picture         =   "frm工伤认定编号选择.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvw工认认定记录 
      Height          =   4245
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   7488
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "认定编号"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "职业病类别"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "伤害部位"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "认定结论"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "伤残等级"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frm工伤认定编号选择"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrCode As String

Public Function ShowME() As String
    On Error Resume Next
    mstrCode = ""
    Me.Show 1
    ShowME = mstrCode
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Dim lvwItem As ListItem
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    
    Me.lvw工认认定记录.ListItems.Clear
    Set nodRowset = mdomOutput.selectSingleNode("DATA").childNodes(6)
    For Each nodRow In nodRowset.childNodes
        Set lvwItem = lvw工认认定记录.ListItems.Add(, "K_" & lvw工认认定记录.ListItems.Count + 1, nodRow.childNodes(0).nodeTypedValue, 1, 1)
        lvwItem.SubItems(1) = nodRow.childNodes(3).nodeTypedValue   '职业病类别
        lvwItem.SubItems(2) = nodRow.childNodes(4).nodeTypedValue   '伤害部位
        lvwItem.SubItems(3) = nodRow.childNodes(5).nodeTypedValue   '认定结论
        lvwItem.SubItems(4) = nodRow.childNodes(6).nodeTypedValue   '伤残等级
    Next
    
    If lvw工认认定记录.ListItems.Count = 0 Then
        Unload Me
        Exit Sub
    ElseIf lvw工认认定记录.ListItems.Count = 1 Then
        lvw工认认定记录.ListItems(1).Selected = True
        lvw工认认定记录.SelectedItem.Selected = True
        Call lvw工认认定记录_DblClick
        Exit Sub
    End If
End Sub

Private Sub lvw工认认定记录_DblClick()
    If lvw工认认定记录.ListItems.Count = 0 Then Exit Sub
    If lvw工认认定记录.SelectedItem Is Nothing Then Exit Sub
    
    mstrCode = lvw工认认定记录.SelectedItem.Text
    Unload Me
End Sub

Private Sub lvw工认认定记录_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call lvw工认认定记录_DblClick
End Sub
