VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm设置本地支持的医保 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置本地支持的医保"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   Icon            =   "frm设置本地支持的医保.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.ImageList imgLvw 
      Left            =   4230
      Top             =   180
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
            Picture         =   "frm设置本地支持的医保.frx":1272
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5640
      TabIndex        =   3
      Top             =   1470
      Width           =   1100
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5640
      TabIndex        =   2
      Top             =   960
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvw所有医保 
      Height          =   3435
      Left            =   240
      TabIndex        =   0
      Top             =   780
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   6059
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgLvw"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "名称"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "险类"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "外挂"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "说明"
         Object.Width           =   4410
      EndProperty
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "请选择本地支持的医保"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1170
      TabIndex        =   1
      Top             =   360
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   390
      Picture         =   "frm设置本地支持的医保.frx":24F4
      Top             =   180
      Width           =   240
   End
End
Attribute VB_Name = "frm设置本地支持的医保"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnStart As Boolean                '是否启动
Private mstrSelect As String                '支持的医保

Private Sub cmd取消_Click()
    Unload Me
End Sub

Private Sub cmd确定_Click()
    Dim intDO As Integer, intCOUNT As Integer
    Dim int外挂 As Integer          '某院不允许同时使用多个外挂医保工具
    
    intCOUNT = lvw所有医保.ListItems.Count
    '检查有多少个外挂医保工具
    For intDO = 1 To intCOUNT
        If lvw所有医保.ListItems(intDO).Checked Then
            If lvw所有医保.ListItems(intDO).Tag = 1 Then
                int外挂 = int外挂 + 1
            End If
        End If
    Next
    If int外挂 > 1 Then
        MsgBox "目前不支持同时使用多个外挂医保！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '组织成串
    For intDO = 1 To intCOUNT
        If lvw所有医保.ListItems(intDO).Checked Then mstrSelect = mstrSelect & "," & Mid(lvw所有医保.ListItems(intDO).Key, 3)
    Next
    If mstrSelect <> "" Then mstrSelect = Mid(mstrSelect, 2)
    Call SaveSetting("ZLSOFT", "公共全局", "本地支持的医保", mstrSelect)
    
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
    Dim arrSelect
    Dim lvwItem As ListItem
    Dim intDO As Integer, intCOUNT As Integer
    Dim rsTemp As New ADODB.Recordset
    mstrSelect = GetSetting("ZLSOFT", "公共全局", "本地支持的医保", "")
    
    '说明：选择本地支持的险类
    gstrSQL = " Select A.序号,A.名称,A.说明,Nvl(A.外挂,0) AS 外挂" & _
              " From 保险类别 A Where Nvl(是否禁止,0)=0" & _
              " Order By A.序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取已安装并允许使用的医保接口")
    If rsTemp.RecordCount = 0 Then
        MsgBox "由于没有安装任何医保接口，您无法为本地选择支持的医保！", vbInformation, gstrSysName
        Exit Sub
    Else
'        rsTemp.Filter = "启用=1"
'        If rsTemp.RecordCount = 0 Then
'            MsgBox "由于没有启用医保接口，您无法为本地选择支持的医保！", vbInformation, gstrSysName
'            rsTemp.Filter = 0
'            Exit Sub
'        End If
    End If
    
    Me.lvw所有医保.ListItems.Clear
    With rsTemp
        Do While Not .EOF
            Set lvwItem = lvw所有医保.ListItems.Add(, "K_" & !序号, !名称, , 1)
            lvwItem.SubItems(1) = !序号
            lvwItem.SubItems(2) = IIf(!外挂 = 1, "是", "否")
            lvwItem.SubItems(3) = Nvl(!说明)
            lvwItem.Tag = !外挂
            .MoveNext
        Loop
        lvw所有医保.ListItems(1).Selected = True
    End With
    
    '显示本地支持的医保
    On Error Resume Next
    arrSelect = Split(mstrSelect, ",")
    intCOUNT = UBound(arrSelect)
    For intDO = 0 To intCOUNT
        lvw所有医保.ListItems("K_" & arrSelect(intDO)).Checked = True
    Next
    
    mstrSelect = ""
    mblnStart = True
End Sub
