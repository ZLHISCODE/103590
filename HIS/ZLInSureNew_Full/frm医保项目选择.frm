VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm医保项目选择 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "医保项目选择"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   Icon            =   "frm医保项目选择.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5100
      TabIndex        =   3
      Top             =   3420
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3900
      TabIndex        =   2
      Top             =   3420
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvwSelect 
      Height          =   2355
      Left            =   210
      TabIndex        =   1
      Top             =   900
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   4154
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "编码"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "名称"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "说明"
         Object.Width           =   4410
      EndProperty
   End
   Begin MSComctlLib.ImageList imglvw 
      Left            =   60
      Top             =   1350
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
            Picture         =   "frm医保项目选择.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lab项目信息 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   510
      Width           =   6075
   End
   Begin VB.Label LblNote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "当前输入的项目设置了多个医保编码，请选择"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3600
   End
End
Attribute VB_Name = "frm医保项目选择"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstr项目编码 As String
Private mrsData As New ADODB.Recordset
Private mlng收费细目ID As Long

Private Sub cmdCancel_Click()
    mlng收费细目ID = 0
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim intDO As Integer, intCOUNT As Integer
    Dim intSelect As Integer
    
    mstr项目编码 = ""
    If lvwSelect.ListItems.Count = 0 Then Exit Sub
    If lvwSelect.SelectedItem Is Nothing Then Exit Sub
    
    mstr项目编码 = lvwSelect.SelectedItem
    
    Unload Me
End Sub

Public Function ShowSelect(ByVal rsData As ADODB.Recordset, ByVal lng收费细目ID As Long) As String
    
    mstr项目编码 = ""
    Set mrsData = rsData
    mlng收费细目ID = lng收费细目ID
    Me.Show 1
    ShowSelect = mstr项目编码
End Function

Private Sub Form_Load()
    Dim lvwItem As ListItem
    Dim rsSfxm As New ADODB.Recordset
    
    lvwSelect.ListItems.Clear
    Set lvwSelect.SmallIcons = imglvw
    
    With mrsData
        Do While Not .EOF
            Set lvwItem = lvwSelect.ListItems.Add(, "K_" & !编码, !编码, , 1)
            lvwItem.SubItems(1) = !名称
            lvwItem.SubItems(2) = Nvl(!说明)
            .MoveNext
        Loop
    End With
    
    gstrSQL = "Select 编码||' '||名称 as 信息 From 收费细目 Where ID=[1]"
    Set rsSfxm = zlDatabase.OpenSQLRecord(gstrSQL, "收费细目", mlng收费细目ID)
    If rsSfxm.RecordCount > 0 Then
        lab项目信息.Caption = rsSfxm!信息
    End If
End Sub

Private Sub lvwSelect_DblClick()
    Call cmdOK_Click
End Sub

Private Sub lvwSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call cmdOK_Click
End Sub
