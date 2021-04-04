VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm选择当前医保 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "选择当前医保"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5445
   Icon            =   "frm选择当前医保.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2820
      TabIndex        =   1
      Top             =   2580
      Width           =   1100
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4140
      TabIndex        =   2
      Top             =   2580
      Width           =   1100
   End
   Begin MSComctlLib.ImageList imgLvw 
      Left            =   3990
      Top             =   0
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
            Picture         =   "frm选择当前医保.frx":1272
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvw所有医保 
      Height          =   2475
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   4366
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgLvw"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "序号"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "名称"
         Object.Width           =   5292
      EndProperty
   End
End
Attribute VB_Name = "frm选择当前医保"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnStart As Boolean
Private mint险类 As Integer
Private mstrSelect As String
'从本地支持的医保接口中进行选择，用于门诊收费与住院登记等窗口程序

Private Sub cmd取消_Click()
    mint险类 = 0
    Unload Me
    Exit Sub
End Sub

Private Sub cmd确定_Click()
    If lvw所有医保.SelectedItem Is Nothing Then Exit Sub
    mint险类 = Val(lvw所有医保.SelectedItem)
    
    Unload Me
    Exit Sub
End Sub

Private Sub Form_Activate()
    If Not mblnStart Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Dim lvwItem As ListItem
    Dim rsTemp As New ADODB.Recordset
    
    mblnStart = False
    mstrSelect = GetSetting("ZLSOFT", "公共全局", "本地支持的医保", "")
    If mstrSelect = "" Then Exit Sub
    
    '说明：选择本地支持的险类
    gstrSQL = " Select A.序号,A.名称,A.说明" & _
              " From 保险类别 A,Table(Cast(f_num2list([1]) As Zltools.t_Numlist)) B" & _
              " Where A.序号 = B.Column_Value And A.医保包 Is NULL" & _
              " UNION " & _
              " Select DISTINCT A.序号,A.名称,A.说明" & _
              " From 保险类别 B,保险类别 A,Table(Cast(f_num2list([1]) As Zltools.t_Numlist)) C" & _
              " Where B.医保包=A.医保部件 And B.序号 = C.Column_Value" & _
              " Order By 序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取本地支持的险类", mstrSelect)
    If rsTemp.RecordCount = 0 Then
        MsgBox "还未设置本地支持的医保！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    lvw所有医保.ListItems.Clear
    With rsTemp
        Do While Not .EOF
            Set lvwItem = lvw所有医保.ListItems.Add(, "K_" & !序号, !序号, , 1)
            lvwItem.SubItems(1) = !名称
            .MoveNext
        Loop
        lvw所有医保.ListItems(1).Selected = True
    End With
    
    mblnStart = True
End Sub

Public Function ShowSelect() As Integer
    '显示窗口
    Dim rtn         As Long
    rtn = SetWindowPos(Me.hwnd, -1, CurrentX, CurrentY, 0, 0, 3)
    Me.Show 1
    ShowSelect = mint险类
End Function

Private Sub lvw所有医保_DblClick()
    Call cmd确定_Click
End Sub

Private Sub lvw所有医保_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    Call cmd确定_Click
End Sub
