VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMedicalStationRptPrintSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "打印设置"
   ClientHeight    =   4395
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8040
   Icon            =   "frmMedicalStationRptPrintSet.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.ImageList ils16 
      Left            =   7440
      Top             =   1875
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
            Picture         =   "frmMedicalStationRptPrintSet.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   4245
      Left            =   90
      TabIndex        =   2
      Top             =   75
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   7488
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils16"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "编号"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "名称"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "说明"
         Object.Width           =   4233
      EndProperty
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "退出(&X)"
      Height          =   350
      Left            =   6840
      TabIndex        =   1
      Top             =   495
      Width           =   1100
   End
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "设置(&S)"
      Height          =   350
      Left            =   6840
      TabIndex        =   0
      Top             =   105
      Width           =   1100
   End
End
Attribute VB_Name = "frmMedicalStationRptPrintSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Function ShowEdit(ByVal frmMain As Object, ByVal lng登记id As Long, ByVal lng病人id As Long) As Boolean
    
    Dim rs As New ADODB.Recordset
    Dim objItem As ListItem
    
    gstrSQL = "Select DISTINCT 'ZLCISBILL'||Trim(To_Char(D.编号,'00000'))||'-2' As 报表编号,D.名称,D.说明 " & _
                "From 体检项目清单 A,体检项目医嘱 B,诊疗单据应用 C,病历文件目录 D " & _
                "Where A.登记id = [1] And A.ID = B.清单ID " & _
                "AND C.诊疗项目ID=A.诊疗项目ID AND C.应用场合=4 " & _
                "AND D.ID=C.病历文件ID"
                
    If lng病人id > 0 Then
        gstrSQL = gstrSQL & " And B.病人id=[2]"
    End If
    
    gstrSQL = gstrSQL & " Union All "
    
    gstrSQL = gstrSQL & " Select 编号 As 报表编号,名称,说明 From zlReports Where 编号='ZL1_BILL_1861_2_1'"
    gstrSQL = gstrSQL & " Union All "
    gstrSQL = gstrSQL & " Select 编号 As 报表编号,名称,说明 From zlReports Where 编号='ZL1_BILL_1861_2'"
    gstrSQL = gstrSQL & " Union All "
    gstrSQL = gstrSQL & " Select 编号 As 报表编号,名称,说明 From zlReports Where 编号='ZL1_BILL_1861_2_2'"
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng登记id, lng病人id)
    If rs.BOF Then Exit Function
    
    lvw.ListItems.Clear
    
    Do While Not rs.EOF
        
        Set objItem = lvw.ListItems.Add(, rs("报表编号").Value, rs("报表编号").Value, 1, 1)
        objItem.SubItems(1) = zlCommFun.NVL(rs("名称").Value)
        objItem.SubItems(2) = zlCommFun.NVL(rs("说明").Value)
        
        rs.MoveNext
    Loop
    
    If lvw.ListItems.Count = 0 Then Exit Function
    
    Me.Show 1, frmMain

End Function

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrintSet_Click()
    
    If lvw.SelectedItem Is Nothing Then Exit Sub
    Call ReportPrintSet(gcnOracle, glngSys, lvw.SelectedItem.Key, Me)
    
End Sub


