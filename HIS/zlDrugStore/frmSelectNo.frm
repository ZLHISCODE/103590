VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelectNo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "选择单据"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7815
   Icon            =   "frmSelectNo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   7815
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancle 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6720
      TabIndex        =   3
      Top             =   4080
      Width           =   975
   End
   Begin MSComctlLib.ListView lvwList 
      Height          =   3615
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   7570
      _ExtentX        =   13361
      _ExtentY        =   6376
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "类别"
         Text            =   "类别"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "NO"
         Text            =   "NO"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "姓名"
         Text            =   "姓名"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "性别"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "年龄"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "日期"
         Text            =   "日期"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "单据号"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5400
      TabIndex        =   0
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label lblmsg 
      Caption         =   "请选择需要签到的单据："
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmSelectNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsData As ADODB.Recordset

Public Sub ShowMe(ByRef rsData As Recordset, ByVal frmParent As Form, ByVal strText As String)
    Dim i As Integer
    Dim strType As String
    Dim blnDo As Boolean
    Dim listItem As listItem
    With rsData
        .Filter = "NO='" & strText & "'"
        blnDo = True
        If .RecordCount > 1 Then blnDo = False
        .Filter = ""
    
        Do While Not .EOF
            If !单据 = 8 Then
                strType = "收费单"
            ElseIf !单据 = 9 Then
                strType = "记账单"
            Else
                strType = "记账表"
            End If
            i = i + 1
            Set listItem = Me.lvwList.ListItems.Add(, "k" & i, strType)
            listItem.SubItems(1) = !NO
            listItem.SubItems(2) = !姓名
            listItem.SubItems(3) = NVL(!性别)
            listItem.SubItems(4) = NVL(!年龄)
            listItem.SubItems(5) = NVL(!填制日期)
            listItem.SubItems(6) = !单据
            
            If !NO = strText And blnDo Then listItem.Checked = True
            .MoveNext
        Loop
    End With
    
    Set rsData = Nothing
    
    Me.Show 1, frmParent
    
    Set rsData = mrsData
    Set mrsData = Nothing
End Sub

Private Sub cmdCancle_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim i As Integer
    
    Set mrsData = New ADODB.Recordset
    With mrsData
        If .State = 1 Then .Close
           
        .Fields.Append "单据", adSmallInt
        .Fields.Append "NO", adVarChar, 20
           
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        For i = 1 To Me.lvwList.ListItems.Count
            If Me.lvwList.ListItems(i).Checked = True Then
                .AddNew
                !单据 = Me.lvwList.ListItems(i).SubItems(6)
                !NO = Me.lvwList.ListItems(i).SubItems(1)
                .Update
            End If
        Next
        
    End With
    Unload Me
End Sub
