VERSION 5.00
Begin VB.Form frmStuffParaSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "参数设置"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4860
   Icon            =   "frmStuffParaSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   4860
   StartUpPosition =   2  '屏幕中心
   Begin VB.ComboBox cboSendDept 
      Height          =   300
      ItemData        =   "frmStuffParaSet.frx":6852
      Left            =   1080
      List            =   "frmStuffParaSet.frx":6854
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   112
      Width           =   3375
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3840
      TabIndex        =   5
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2640
      TabIndex        =   4
      Top             =   3600
      Width           =   975
   End
   Begin VB.Frame fraOther 
      Caption         =   "智能卡及其他设备定义"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   4575
      Begin VB.CommandButton cmdDeviceSetup 
         Caption         =   "设备配置"
         Height          =   350
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame fraUnit 
      Caption         =   "缺省单位"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   4575
      Begin VB.OptionButton optUnit 
         Caption         =   "包装单位"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   13
         Top             =   300
         Width           =   1215
      End
      Begin VB.OptionButton optUnit 
         Caption         =   "散装单位"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.Frame fraType 
      Caption         =   "业务类型"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4575
      Begin VB.ComboBox cboNO 
         Height          =   300
         ItemData        =   "frmStuffParaSet.frx":6856
         Left            =   960
         List            =   "frmStuffParaSet.frx":6858
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   700
         Width           =   3375
      End
      Begin VB.CheckBox chkType 
         Caption         =   "记帐表"
         Height          =   255
         Index           =   2
         Left            =   3480
         TabIndex        =   9
         Top             =   300
         Width           =   855
      End
      Begin VB.CheckBox chkType 
         Caption         =   "记帐单"
         Height          =   255
         Index           =   1
         Left            =   2220
         TabIndex        =   8
         Top             =   300
         Width           =   855
      End
      Begin VB.CheckBox chkType 
         Caption         =   "收费单"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   7
         Top             =   300
         Width           =   855
      End
      Begin VB.Label lblNo 
         Caption         =   "收费单据"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   723
         Width           =   735
      End
      Begin VB.Label lblType 
         Caption         =   "单据类型"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   855
      End
   End
   Begin VB.Label lblSendDept 
      Caption         =   "发料部门"
      Height          =   225
      Left            =   240
      TabIndex        =   14
      Top             =   150
      Width           =   735
   End
End
Attribute VB_Name = "frmStuffParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModule As Long  '模块号
Private mstrPrivs As String '权限串
Private mblnOk As Boolean   '参数设置是否成功

Private Sub cmdCancle_Click()
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDeviceSetup_Click()
    Call FS.DeviceSetup(Me, glngSys, mlngModule)
End Sub

Private Sub CmdSave_Click()
    Dim str业务类型 As String
    
    str业务类型 = IIf(chkType(0).Value = 1, "24", "0")
    str业务类型 = str业务类型 & IIf(chkType(1).Value = 1, ",25", ",0")
    str业务类型 = str业务类型 & IIf(chkType(2).Value = 1, ",26", ",0")
    
    On Error GoTo ErrHandle
    Call zlDatabase.SetPara("查询业务类型", str业务类型, glngSys, mlngModule)
    Call zlDatabase.SetPara("卫材单位", IIf(optUnit(1).Value = True, 1, 0), glngSys, mlngModule)
    Call zlDatabase.SetPara("收费处方显示方式", cboNO.ListIndex, glngSys, mlngModule)
    
    Call zlDatabase.SetPara("发料科室", cboSendDept.ItemData(cboSendDept.ListIndex), glngSys, mlngModule)
    Unload Me
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim strReg As String
    Dim arrStr As Variant
    Dim i As Integer
    Dim blnSetPara As Boolean
    Dim lng发料部门ID As Long

    With cboNO
        .Clear
        .AddItem "1-显示所有的处方"
        .AddItem "2-仅显示已收费处方"
        .AddItem "3-仅显示未收费处方"
        .ListIndex = 0
    End With
    
    blnSetPara = (InStr(1, mstrPrivs, "参数设置") > 0)
    
    strReg = Val(zlDatabase.GetPara("收费处方显示方式", glngSys, mlngModule, 0, Array(LblNo, cboNO), blnSetPara))
    If Val(strReg) >= 0 And strReg <= 2 Then
        cboNO.ListIndex = Val(strReg)
    Else
        cboNO.ListIndex = 0
    End If
    
    strReg = Val(zlDatabase.GetPara("卫材单位", glngSys, mlngModule, "0", Array(optUnit(0), optUnit(1), fraUnit), blnSetPara))
    optUnit(0).Value = False
    optUnit(1).Value = False
    If Val(strReg) = 0 Then
        optUnit(0).Value = True
    Else
        optUnit(1).Value = True
    End If
    
    strReg = Trim(zlDatabase.GetPara("查询业务类型", glngSys, mlngModule, "", Array(lblType, chkType(0), chkType(1), chkType(2), fraType), blnSetPara))
    If strReg = "" Then strReg = "24,25,26"
    arrStr = Split(strReg & "," & "," & ",", ",")
    For i = 0 To UBound(arrStr)
        If i > 2 Then Exit For
        chkType(i).Value = IIf(Val(arrStr(i)) > 0, 1, 0)
    Next
    
    lng发料部门ID = Val(zlDatabase.GetPara("发料科室", glngSys, mlngModule, "0", Array(lblSendDept, cboSendDept), blnSetPara))
    Call LoadDept(lng发料部门ID)
    
    
End Sub

Public Function ShowSetPara(ByVal frmMain As Form, ByVal lngModule As Long, ByVal strPrivs As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:设置参数入口
    '参数:
    '返回:设置成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs
    '本地参数设置
    Me.Show 1, frmMain
    ShowSetPara = mblnOk
End Function

Private Sub LoadDept(ByVal lng发料部门ID As Long)
    Dim rsTemp As Recordset
    
    Set rsTemp = Stuff_GetDept(mstrPrivs)
    
    '装入发料部门数据
    With cboSendDept
        .Clear
        Do While Not rsTemp.EOF
            .AddItem rsTemp!名称
            .ItemData(.NewIndex) = rsTemp!Id
            If rsTemp!Id = lng发料部门ID Then
                .ListIndex = .NewIndex
            End If
            rsTemp.MoveNext
        Loop
        If .ListIndex = -1 Then .ListIndex = 0
    End With
End Sub
