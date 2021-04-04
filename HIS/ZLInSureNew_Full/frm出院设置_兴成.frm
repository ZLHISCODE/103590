VERSION 5.00
Begin VB.Form frm出院设置_兴成 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "医保病人出院设置"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5445
   Icon            =   "frm出院设置_兴成.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cbo级别 
      Height          =   300
      Left            =   1350
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2550
      Width           =   3660
   End
   Begin VB.TextBox TxtEdit 
      Height          =   300
      Left            =   1350
      MaxLength       =   20
      TabIndex        =   8
      Top             =   2160
      Width           =   3660
   End
   Begin VB.ComboBox cbo入院类别 
      Height          =   300
      Left            =   1350
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   945
      Width           =   3660
   End
   Begin VB.ComboBox cbo住院类别 
      Height          =   300
      Left            =   1350
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1365
      Width           =   3660
   End
   Begin VB.ComboBox cbo出院类别 
      Height          =   300
      Left            =   1350
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1770
      Width           =   3660
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4095
      TabIndex        =   14
      Top             =   3195
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2910
      TabIndex        =   13
      Top             =   3195
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   45
      Index           =   0
      Left            =   60
      TabIndex        =   12
      Top             =   720
      Width           =   7665
   End
   Begin VB.Frame fra 
      Height          =   45
      Index           =   1
      Left            =   -1935
      TabIndex        =   11
      Top             =   3000
      Width           =   7665
   End
   Begin VB.Label lblinfo 
      AutoSize        =   -1  'True
      Caption         =   "异地医院级别"
      Height          =   180
      Index           =   3
      Left            =   255
      TabIndex        =   9
      Top             =   2610
      Width           =   1080
   End
   Begin VB.Label lblinfo 
      AutoSize        =   -1  'True
      Caption         =   "异地医院名称"
      Height          =   180
      Index           =   2
      Left            =   255
      TabIndex        =   7
      Top             =   2220
      Width           =   1080
   End
   Begin VB.Label lblInfor 
      Caption         =   "入院类别"
      Height          =   210
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   990
      Width           =   735
   End
   Begin VB.Label lblinfo 
      AutoSize        =   -1  'True
      Caption         =   "住院类别"
      Height          =   180
      Index           =   0
      Left            =   615
      TabIndex        =   3
      Top             =   1425
      Width           =   720
   End
   Begin VB.Label lblinfo 
      AutoSize        =   -1  'True
      Caption         =   "出院类别"
      Height          =   180
      Index           =   1
      Left            =   615
      TabIndex        =   5
      Top             =   1830
      Width           =   720
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   150
      Picture         =   "frm出院设置_兴成.frx":0E42
      Top             =   165
      Width           =   480
   End
   Begin VB.Label lblNote 
      Caption         =   "    设置出院病人的入院类别、住院类别及出院类别,当出院类别为转外地时，则需设置异地医院名称及级别."
      Height          =   390
      Left            =   885
      TabIndex        =   0
      Top             =   240
      Width           =   4500
   End
End
Attribute VB_Name = "frm出院设置_兴成"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mblnOK As Boolean



Private Function IsValid() As Boolean
    '判断错误
    IsValid = False
    If zlCommFun.StrIsValid(txtEdit.Text, txtEdit.MaxLength) = False Then
        zlControl.TxtSelAll txtEdit
        txtEdit.SetFocus
        Exit Function
    End If
    IsValid = True
End Function

Private Sub cbo出院类别_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab

End Sub



Private Sub cbo级别_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub



Private Sub cbo入院类别_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo住院类别_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    If IsValid = False Then Exit Sub
    '过程参数:
    '       住院类别_IN,出院类别_IN,异地医院名称_IN,异地医院级别_IN
    
    gstrSQL = "ZL_医保病人出院登记_UPDATE("
    gstrSQL = gstrSQL & mlng病人ID & ","
    gstrSQL = gstrSQL & "'" & cbo住院类别.ItemData(cbo住院类别.ListIndex) & "',"
    gstrSQL = gstrSQL & "'" & cbo出院类别.ItemData(cbo出院类别.ListIndex) & "',"
    If cbo出院类别.ItemData(cbo出院类别.ListIndex) = 3 Then
        gstrSQL = gstrSQL & "'" & txtEdit.Text & "',"
        gstrSQL = gstrSQL & "'" & cbo出院类别.ItemData(cbo出院类别.ListIndex) & "')"
    Else
        gstrSQL = gstrSQL & "NULL,"
        gstrSQL = gstrSQL & "NULL)"
    End If
    ExecuteProcedure_兴成 Me.Caption
    gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_兴成核工业 & ",'人员身份','" & cbo入院类别.ItemData(cbo入院类别.ListIndex) & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存入院类别")
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_Load()

    Call InitData
End Sub

Private Sub InitData()
    Dim rsTemp As New ADODB.Recordset
    Dim str入院类别 As String
    Dim i As Long
    Me.cbo入院类别.Clear
    Me.cbo住院类别.Clear
    Me.cbo出院类别.Clear
    With Me.cbo入院类别
        .AddItem "1-正常入院"
        .ItemData(.NewIndex) = 1
        .ListIndex = 0
        .ListIndex = .NewIndex
        .AddItem "2-市内转入"
        .ItemData(.NewIndex) = 2
        .AddItem "3-市外转入"
        .ItemData(.NewIndex) = 3
        .AddItem "4-因慢性病加重第一次住院"
        .ItemData(.NewIndex) = 4
    End With
    With Me.cbo住院类别
        .AddItem "0-正常住院"
        .ItemData(.NewIndex) = 0
        .ListIndex = .NewIndex
        .AddItem "1-紧急抢救"
        .ItemData(.NewIndex) = 1
    End With
    With Me.cbo出院类别
        .AddItem "1-正常出院"
        .ItemData(.NewIndex) = 1
        .ListIndex = .NewIndex
        .AddItem "2-转往市内"
        .ItemData(.NewIndex) = 2
        .AddItem "3-转往市外"
        .ItemData(.NewIndex) = 3
    End With
    
    '入院类别确认
    gstrSQL = "Select 病人id,人员身份 From 保险帐户 where 病人id=" & mlng病人ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    If rsTemp.EOF = False Then
        For i = 0 To Me.cbo入院类别.ListCount - 1
            If Me.cbo入院类别.ItemData(i) = Val(Nvl(rsTemp!人员身份)) Then
                Me.cbo入院类别.ListIndex = i: Exit For
            End If
            
        Next
    End If
        
    '出院类别确认
    gstrSQL = "Select AF17,AF18,AF19 as 住院类别,AF20 as 出院类别 From 医保病人附加信息 where 病人id=" & mlng病人ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    With Me.cbo住院类别
        For i = 0 To .ListCount - 1
            If .ItemData(i) = Val(Nvl(rsTemp!住院类别)) Then
                .ListIndex = i: Exit For
            End If
        Next
    End With
    With Me.cbo出院类别
        For i = 0 To .ListCount - 1
            If .ItemData(i) = Val(Nvl(rsTemp!出院类别)) Then
                .ListIndex = i: Exit For
            End If
        Next
    End With
End Sub


Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab

End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtEdit, KeyAscii, m文本式
End Sub
Private Sub cbo出院类别_Click()
     
    If cbo出院类别.ItemData(cbo出院类别.ListIndex) <> 3 Then
        Me.cbo级别.Clear
        With Me.cbo级别
            .AddItem "11: 转市内一级医院"
            .ItemData(.NewIndex) = 11
            .ListIndex = .NewIndex
            .AddItem "12: 转市内二级医院"
            .ItemData(.NewIndex) = 12
            .AddItem "13: 转市内三级医院"
            .ItemData(.NewIndex) = 13
            .AddItem "14: 转市外一级医院"
            .ItemData(.NewIndex) = 14
            .AddItem "15: 转市外二级医院"
            .ItemData(.NewIndex) = 15
            .AddItem "16: 转市外三级医院"
            .ItemData(.NewIndex) = 16
            .AddItem "17: 转省外一级医院"
            .ItemData(.NewIndex) = 17
            .AddItem "18: 转省外二级医院"
            .ItemData(.NewIndex) = 18
            .AddItem "19: 转省外三级医院"
            .ItemData(.NewIndex) = 19
        End With
        Exit Sub
    Else
        Me.cbo级别.Clear
'        With Me.cbo级别
'            .AddItem "01: 一级医院"
'            .ItemData(.NewIndex) = 1
'            .AddItem "02: 二级医院"
'            .ItemData(.NewIndex) = 2
'            .AddItem "03: 三级医院"
'            .ItemData(.NewIndex) = 3
'        End With
    End If
End Sub
Public Function ShowCard(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    Me.Show vbModal
    ShowCard = mblnOK
End Function
