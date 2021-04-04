VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPACSImageDeviceSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "影像设备设置"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8340
   FillColor       =   &H00FF0000&
   Icon            =   "frmPACSImageDeviceSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7170
      TabIndex        =   2
      Top             =   5340
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   6030
      TabIndex        =   1
      Top             =   5340
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvwImageDevice 
      Height          =   4935
      Left            =   30
      TabIndex        =   0
      Top             =   300
      Width           =   8265
      _ExtentX        =   14579
      _ExtentY        =   8705
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "设备号"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "设备名"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "IP地址"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "端口号"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "本地AE"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "远程设备AE"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "默认设备"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "双击选中一个设备作为默认设备："
      ForeColor       =   &H8000000D&
      Height          =   180
      Left            =   90
      TabIndex        =   3
      Top             =   75
      Width           =   2700
   End
End
Attribute VB_Name = "frmPACSImageDeviceSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstrDeviceName As String
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "默认影像设备", mstrDeviceName
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim objItem As ListItem
    
    strSQL = "select 设备号 , 设备名, IP地址,端口号,本地AE,设备AE from 影像设备目录 where 类型 = 4 "
    
    On Error GoTo ErrOther
    
    '读入现有影像设备
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Me.lvwImageDevice.ListItems.Clear
    Do Until rsTmp.EOF
        With Me.lvwImageDevice.ListItems
            Set objItem = .Add(, "_" & rsTmp("设备号"), rsTmp("设备号"))
            objItem.SubItems(1) = rsTmp("设备名")
            objItem.SubItems(2) = rsTmp("IP地址")
            objItem.SubItems(3) = Nvl(rsTmp("端口号"))
            objItem.SubItems(4) = Nvl(rsTmp("本地AE"))
            objItem.SubItems(5) = Nvl(rsTmp("设备AE"))
        End With
        rsTmp.MoveNext
    Loop
    
    '设置设备默认值
    mstrDeviceName = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "默认影像设备", "")
    If mstrDeviceName <> "" Then
        Set objItem = Me.lvwImageDevice.FindItem(Mid(mstrDeviceName, 2))
        If Not objItem Is Nothing Then
            objItem.SubItems(6) = "√"
        Else
'            If Me.lvwImageDevice.ListItems.Count > 0 Then Me.lvwImageDevice.ListItems(1).SubItems(6) = "√"
        End If
    Else
'        If Me.lvwImageDevice.ListItems.Count > 0 Then Me.lvwImageDevice.ListItems(1).SubItems(6) = "√"
    End If
    
    Exit Sub
    
ErrOther:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Sub subSelectDefault()
    '------------------------------------------------
    '功能：选择一个默认设备
    '参数：
    '返回：无
    '上级函数或过程：lvwImageDevice_DblClick；
    '下级函数或过程：无
    '引用的外部参数：mstrDeviceName
    '编制人：曾超
    '------------------------------------------------
    Dim i As Integer
    
    If Me.lvwImageDevice.SelectedItem Is Nothing Then Exit Sub
    
    With Me.lvwImageDevice
        For i = 1 To .ListItems.Count
            .ListItems(i).SubItems(6) = ""
        Next
    End With
    Me.lvwImageDevice.SelectedItem.SubItems(6) = "√"
    mstrDeviceName = Me.lvwImageDevice.SelectedItem.Key
End Sub

Private Sub lvwImageDevice_DblClick()
    subSelectDefault
End Sub
