VERSION 5.00
Begin VB.Form frm就诊类型修改 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "请选择结算类型"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   Icon            =   "frm就诊类型修改.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txt承担比例 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Left            =   1830
      MaxLength       =   3
      TabIndex        =   3
      Top             =   1080
      Width           =   2025
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   30
      TabIndex        =   7
      Top             =   1590
      Width           =   4425
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1770
      TabIndex        =   4
      Top             =   1800
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3015
      TabIndex        =   5
      Top             =   1800
      Width           =   1100
   End
   Begin VB.ComboBox cbo医疗类别 
      Height          =   300
      Left            =   1830
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   690
      Width           =   2025
   End
   Begin VB.Label lbl承担比例 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "承担比例"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1020
      TabIndex        =   2
      Top             =   1140
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   150
      Picture         =   "frm就诊类型修改.frx":000C
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "请为该病人选择结算类型："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1020
      TabIndex        =   6
      Top             =   330
      Width           =   2160
   End
   Begin VB.Label lbl医疗类别 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "医疗类别"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1020
      TabIndex        =   0
      Top             =   750
      Width           =   720
   End
End
Attribute VB_Name = "frm就诊类型修改"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng病人ID As Long

Private Sub cbo医疗类别_Click()
    Me.txt承担比例.Enabled = False
    If cbo医疗类别.ItemData(cbo医疗类别.ListIndex) = 22 Then
        '交通事故
        Me.txt承担比例.Enabled = True
        Me.txt承担比例.SetFocus
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Me.cbo医疗类别.ItemData(Me.cbo医疗类别.ListIndex) = 22 Then
        If Val(txt承担比例.Text) < 0 Then
            MsgBox "承担比例不能小于零！", vbInformation, gstrSysName
            txt承担比例.SetFocus
            Exit Sub
        End If
        If Val(txt承担比例.Text) > 100 Then
            MsgBox "承担比例不能大于一百！", vbInformation, gstrSysName
            txt承担比例.SetFocus
            Exit Sub
        End If
    End If
    
    gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_慈溪农医 & ",'业务类型','''" & Me.cbo医疗类别.ItemData(Me.cbo医疗类别.ListIndex) & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存业务类型")
    
    Unload Me
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Public Sub ShowME(ByVal lng病人ID As Long)
    mlng病人ID = lng病人ID
    Me.Show 1
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    
    With cbo医疗类别
        .AddItem "普通住院"
        .ItemData(.NewIndex) = 21
        .AddItem "交通事故"
        .ItemData(.NewIndex) = 22
        .AddItem "大病救助"
        .ItemData(.NewIndex) = 23
        .AddItem "难产"
        .ItemData(.NewIndex) = 24
        .AddItem "其他"
        .ItemData(.NewIndex) = 25
        .ListIndex = 0
    End With
    
    '提取该病人当前的结算类型
    gstrSQL = "Select Nvl(业务类型,21) AS 结算类型 From 保险帐户 Where 险类=[1] And 病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人当前结算类型", TYPE_慈溪农医, mlng病人ID)
    If rsTemp.RecordCount <> 0 Then
        cbo医疗类别.ListIndex = (rsTemp!结算类型 - 21)
    End If
End Sub
