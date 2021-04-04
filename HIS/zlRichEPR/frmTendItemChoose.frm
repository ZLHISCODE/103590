VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTendItemChoose 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "护理项目选择"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5865
   Icon            =   "frmTendItemChoose.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.ImageList imgLvw 
      Left            =   2640
      Top             =   1710
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
            Picture         =   "frmTendItemChoose.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4500
      TabIndex        =   2
      Top             =   3480
      Width           =   1100
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3210
      TabIndex        =   1
      Top             =   3480
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   3345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   5900
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgLvw"
      SmallIcons      =   "imgLvw"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "项目序号"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "项目名称"
         Object.Width           =   4410
      EndProperty
   End
End
Attribute VB_Name = "frmTendItemChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mlng项目序号 As Long
Private mstr项目名称 As String
Private mstrSelItems As String
Private mrsItems As New ADODB.Recordset

Public Function ShowSelect(ByVal strSelItems As String, ByVal byt护理等级 As Integer, ByVal int婴儿 As Integer, ByVal lng科室ID As Long) As String
    On Error Resume Next
    Dim lvwItem As ListItem
    
    mblnOK = False
    mstrSelItems = strSelItems
    '按以前的规则提取项目清单供录入
    gstrSQL = " Select B.项目序号,B.项目名称 " & _
             " From 护理记录项目 B" & _
             " Where B.应用方式<>0 " & IIf(byt护理等级 = -1, "", " And B.护理等级>=[1]") & IIf(int婴儿 = -1, "", " And B.适用病人 IN (0,[2])") & _
             " And (B.适用科室=1 Or (B.适用科室=2 And Exists (Select 1 From 护理适用科室 D Where D.项目序号=B.项目序号 And D.科室id=[3])))" & _
             " Order by B.项目序号"
    Set mrsItems = zlDatabase.OpenSQLRecord(gstrSQL, "提取所有可用的护理项目", byt护理等级, IIf(int婴儿 = 0, 1, 2), lng科室ID)
    If mrsItems.RecordCount = 0 Then
        MsgBox "没有可供添加的项目！", vbInformation, gstrSysName
        Exit Function
    End If
    '将可选择的项目加入控件中
    lvwItems.ListItems.Clear
    With mrsItems
        Do While Not .EOF
            If InStr(1, mstrSelItems, "," & !项目序号 & ",") = 0 Then
                Set lvwItem = lvwItems.ListItems.Add(, "K" & lvwItems.ListItems.Count, !项目序号, , 1)
                lvwItem.SubItems(1) = !项目名称
            End If
            .MoveNext
        Loop
    End With
    If lvwItems.ListItems.Count = 0 Then
        MsgBox "没有可供添加的项目！", vbInformation, gstrSysName
        Unload Me
        Exit Function
    End If
    
    Me.Show 1
    If mblnOK Then ShowSelect = mlng项目序号 & "|" & mstr项目名称
End Function

Private Sub cmd取消_Click()
    Unload Me
End Sub

Private Sub cmd确定_Click()
    If lvwItems.SelectedItem Is Nothing Then Exit Sub
    
    mlng项目序号 = lvwItems.SelectedItem
    mstr项目名称 = lvwItems.SelectedItem.SubItems(1)
    mblnOK = True
    Unload Me
End Sub

Private Sub lvwItems_DblClick()
    Call lvwItems_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub lvwItems_KeyDown(KeyCode As Integer, Shift As Integer)
    If lvwItems.SelectedItem Is Nothing Then Exit Sub
    If KeyCode = vbKeyReturn Then Call cmd确定_Click
End Sub
