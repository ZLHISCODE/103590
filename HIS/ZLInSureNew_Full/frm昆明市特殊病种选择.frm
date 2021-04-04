VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm昆明市特殊病种选择 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "昆明市特殊病种选择"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   Icon            =   "frm昆明市特殊病种选择.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4950
      TabIndex        =   4
      Top             =   3780
      Width           =   1100
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2430
      Top             =   1830
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
            Picture         =   "frm昆明市特殊病种选择.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3690
      TabIndex        =   3
      Top             =   3780
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -180
      TabIndex        =   2
      Top             =   3600
      Width           =   6465
   End
   Begin MSComctlLib.ListView lvwDisease 
      Height          =   3075
      Left            =   150
      TabIndex        =   1
      Top             =   360
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   5424
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "病种编码"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "病种名称"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "统筹结算标准"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "适用范围"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.Label lbl请选择一种特殊疾病 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "请慎重选择一种特殊疾病"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   5925
   End
End
Attribute VB_Name = "frm昆明市特殊病种选择"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrDisease As String

Public Function ShowSelect(ByVal strDisease As String) As String
    On Error Resume Next
    mstrDisease = strDisease
    Me.Show 1
    ShowSelect = mstrDisease
End Function

Private Sub cmd取消_Click()
    mstrDisease = ""
    Unload Me
End Sub

Private Sub Cmd确定_Click()
    If lvwDisease.SelectedItem Is Nothing Then
        MsgBox "请选择一种特殊疾病！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    mstrDisease = lvwDisease.SelectedItem.Text
    Unload Me
End Sub

Private Sub Form_Load()
    Dim arrDisease
    Dim lvwItem As ListItem
    Dim intDO As Integer, intMax As Integer
    Dim rsTemp As New ADODB.Recordset
    
    '打开医保前置机
    With gcnSybase
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        '固定使用该用户、密码和主机字符串
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & "si2000", "his", "his"
        '.Open "Driver={Microsoft ODBC for Oracle};Server=" & "", "zlhis", "his"
        If .State = adStateClosed Then Unload Me: Exit Sub
    End With
      
    mstrDisease = Replace(mstrDisease, " ", "")
    If mstrDisease = "" Then
        Unload Me
        Exit Sub
    End If
   
    '组织成IN串所需，如：'0101','0102'
    mstrDisease = Replace(mstrDisease, "$", "','")
    mstrDisease = "'" & mstrDisease & "'"
    
    '从前置机中提取病种信息
    gstrSQL = "Select DBzbm AS 病种编码,DBzmc AS 病种名称,Tcjsbz AS 统筹结算标准,Syfw AS 适用范围" & _
        " From V_BY02DBZBZ" & _
        " Where DBzbm in (" & mstrDisease & ")"
    'MsgBox gstrSQL
    Call OpenRecordset(rsTemp, "从前置机中提取病种信息", gstrSQL, gcnSybase)
    
    With rsTemp
        Do While Not .EOF
            Set lvwItem = lvwDisease.ListItems.Add(, "K_" & !病种编码, !病种编码, 1, 1)
            lvwItem.SubItems(1) = !病种名称
            lvwItem.SubItems(2) = !统筹结算标准
            lvwItem.SubItems(3) = !适用范围
            .MoveNext
        Loop
    End With
End Sub

Private Sub lvwDisease_DblClick()
    Call Cmd确定_Click
End Sub

Private Sub lvwDisease_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call Cmd确定_Click
End Sub
