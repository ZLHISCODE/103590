VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelOwner 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "所有者"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   Icon            =   "frmSelOwner.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chkNext 
      Caption         =   "其他对象使用相同的所有者"
      Height          =   255
      Left            =   75
      TabIndex        =   4
      Top             =   2040
      Value           =   1  'Checked
      Width           =   2895
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   1335
      Left            =   75
      TabIndex        =   3
      Top             =   645
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   2355
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "所有者"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "对象名"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   3510
      TabIndex        =   1
      Top             =   870
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3510
      TabIndex        =   0
      Top             =   1335
      Width           =   1100
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   180
      Picture         =   "frmSelOwner.frx":014A
      Top             =   105
      Width           =   240
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "下面的数据对象可以从不同的所有者处获得，请选择一个你想要访问的所有者."
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   150
      Width           =   3690
   End
End
Attribute VB_Name = "frmSelOwner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rsObject As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If lvw.SelectedItem Is Nothing Then
        MsgBox "没有选择！", vbInformation, App.Title: Exit Sub
    End If
    gblnOK = True
    Hide
End Sub

Private Sub Form_Activate()
    lvw.SetFocus
End Sub

Private Sub Form_Load()
    Dim i As Integer, objItem As Object
    Dim CurrentDBUser As String
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    gblnOK = False
    strSQL = "Select 所有者 From zlSystems Where 编号=[1]"
    Set rsData = OpenSQLRecord(strSQL, Me.Caption, glngSys)
    If rsData.BOF = False Then
        CurrentDBUser = UCase(Nvl(rsData("所有者").Value))
    End If
    If Not rsObject Is Nothing Then
        rsObject.MoveFirst
        For i = 1 To rsObject.RecordCount
            Set objItem = lvw.ListItems.Add(, , rsObject!Owner)
            objItem.SubItems(1) = rsObject!OBJECT_NAME
            rsObject.MoveNext
        Next
        If lvw.ListItems.count > 0 Then
            For i = 0 To lvw.ListItems.count - 1
                
                If UCase(lvw.ListItems(i + 1).Text) = UCase(CurrentDBUser) Then
                    lvw.ListItems(i + 1).Selected = True
                    Exit For
                End If
            Next
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsObject = Nothing
End Sub

Private Sub lvw_DblClick()
    cmdOK_Click
End Sub
